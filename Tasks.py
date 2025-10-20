import json
import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError
import warnings
import pyxlsb
import csv
import xlwings as xw



warnings.filterwarnings("ignore", category=UserWarning)
import sys


def Process_A14_options(file_path, q):
    q.put(("status", "🔄 Inicializando o processamento dos arquivos..."))

    # Step 1: Load data (This part is unchanged)
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".xlsx", ".xlsm"]:
            df = pd.read_excel(file_path, engine="openpyxl")
        elif ext == ".xls":
            df = pd.read_excel(file_path, engine="xlrd")
        elif ext == ".xlsb":
            df = pd.read_excel(file_path, engine="pyxlsb")
        elif ext == ".csv":
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                sample = f.read(4096)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
                    delimiter = dialect.delimiter
                except csv.Error:
                    delimiter = ";" if ";" in sample else ","
            try:
                df = pd.read_csv(file_path, delimiter=delimiter, encoding="utf-8", engine="python")
            except Exception:
                df = pd.read_csv(file_path, delimiter=delimiter, encoding="latin-1", engine="python")
        else:
            q.put(("status", f"❌ Formato de arquivo não suportado: {ext}"))
            return
    except Exception as e:
        q.put(("status", f"❌ Erro ao ler arquivo: {e}"))
        return

    q.put(("status", f"✅ Arquivo carregado ({len(df)} linhas, {len(df.columns)} colunas)"))

    if 'CODICE_FAMIGLIA' not in df.columns:
        q.put(("status", f"❌ Coluna 'CODICE_FAMIGLIA' não encontrada. Colunas disponíveis: {list(df.columns)}"))
        return

    # Filter for 'PKG' rows
    df['CODICE_FAMIGLIA'] = df['CODICE_FAMIGLIA'].astype(str).str.strip().str.upper()
    df_pkg = df[df['CODICE_FAMIGLIA'] == 'PKG'].copy()

    if df_pkg.empty:
        q.put(("status", "⚠️ Nenhuma linha encontrada com CODICE_FAMIGLIA = 'PKG'."))
        return

    # =============================================================================
    # NEW LOGIC BASED ON YOUR INSTRUCTIONS
    # =============================================================================

    # Step 1: Find all columns with 'CODICE_OPTIONAL' in the name
    optional_cols = [col for col in df_pkg.columns if 'CODICE_OPTIONAL' in col]
    if not optional_cols:
        q.put(("status", "⚠️ Nenhuma coluna contendo 'CODICE_OPTIONAL' encontrada."))
        return

    # Step 2: The first optional column is for the PACK, the rest for CONTEÚDO
    pack_col_name = optional_cols[0]
    conteudo_cols = optional_cols[1:]
    
    q.put(("status", f"✅ Coluna do PACK: '{pack_col_name}'"))
    q.put(("status", f"✅ Colunas do CONTEÚDO: {len(conteudo_cols)} colunas"))

    processed_data = []
    for _, row in df_pkg.iterrows():
        # Step 3: Get the PACK value from the first optional column
        pack_value = row[pack_col_name]
        pack = str(pack_value).strip() if pd.notna(pack_value) else ""
        
        # Step 4: Get and join all CONTEÚDO values from the other optional columns
        conteudo_values = []
        for col in conteudo_cols:
            value = row[col]
            if pd.notna(value):
                str_value = str(value).strip()
                if str_value:
                    conteudo_values.append(str_value)
        
        conteudo = "*" + "*".join(conteudo_values) + "*" if conteudo_values else ""
        
        processed_data.append({'PACK': pack, 'CONTEÚDO': conteudo})

    df_result = pd.DataFrame(processed_data, columns=['PACK', 'CONTEÚDO'])
    q.put(("status", f"📦 {len(df_result)} registros prontos para atualização."))

    # =============================================================================
    # The file writing logic below is correct and remains unchanged.
    # =============================================================================

    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    Base_folder = os.path.join(base_path, "Bases")

    if not os.path.exists(Base_folder):
        q.put(("status", f"❌ Pasta 'Bases' não encontrada: {Base_folder}"))
        return

    for filename in os.listdir(Base_folder):
        if 'BASE' in filename.upper() and not filename.startswith("~") and filename.lower().endswith(('.xlsb', '.xlsx', '.xlsm')):
            file_full_path = os.path.join(Base_folder, filename)
            q.put(("status", f"📁 Atualizando arquivo: {filename}"))
            try:
                wb = xw.Book(file_full_path)
                if 'A14' in [s.name for s in wb.sheets]:
                    ws = wb.sheets['A14']
                else:
                    ws = wb.sheets.add('A14')
                ws.clear_contents()
                ws.range('A1').value = ['PACK', 'CONTEÚDO']
                ws.range('A2').value = df_result.values.tolist()
                ws.autofit()
                wb.save()
                wb.close()
                q.put(("status", f"✅ Planilha 'A14' atualizada em {filename}"))
            except Exception as e:
                q.put(("status", f"❌ Falha ao processar {filename}: {e}"))

    q.put(("status", "🎉 Processamento concluído com sucesso."))


def download_A14(page,url_order,q,username,password) :
        page.goto(url_order)

        q.put(("status", "Realizando login..."))
        q.put(("progress", 15))
       
        page.locator('[name="j_username"]').fill(username)
        page.locator('[name="j_password"]').fill(password)
        page.locator("button[type='submit']").click()
        q.put(("status", "Login realizado com sucesso!"))

        page.get_by_role("link", name="???tabstd???").hover()
    
        page.locator("li.ui-menuitem >> text=Download Table").click(timeout=200000)
        page.locator("[id=\"filter:codtab_label\"]").click()
        page.locator("[id=\"filter:codtab_panel\"]").get_by_text("???tabA14???").click()
        
        with page.expect_download() as download_info:
            page.get_by_role("button", name="Downloads").click()
           
        download = download_info.value

        file_path = f"Dados/A14.xls"
        os.makedirs("Dados", exist_ok=True)
        
        if os.path.exists(file_path):
            os.remove(file_path)
        
        download.save_as(file_path)

        Process_A14_options(file_path,q)

        q.put(("status", f"Relatório A14 salvo como: {file_path}"))
        q.put(("status", "Downloads concluídos."))
        q.put(("progress", 65))





# def download_por_modelo(page,url_oss,q,username,password,Modelos) :
#     page.goto(url_oss, timeout= 600000 );
#     # 3. Login
#     q.put(("status", "Realizando login..."))
#     q.put(("progress", 15))

#     page.locator('[name="USER_NAME"]').fill(username)
#     page.locator('[name="PASSWORD"]').fill(password)
#     page.locator(".signin").click()
#     q.put(("status", "Login realizado com sucesso!"))
    

#     for key, value in Modelos.items(): 
#         q.put(("status", f"Processando modelo {key}"))
#         if key == '611' :    #this line is for us to jump the 611 models and its report is failling causing crash, i could just add try andf catch to prevent crash.
#             continue

#         page.locator(".shellInstance").click()
#         print(key,value)

#         page.locator("#sequencer_ui_instances").select_option(value)
#         page.get_by_role("link", name="Editor de programação").click(timeout=500000)

#         q.put(("status", "Aguardando carregar pagina de relatório do modelo {key}"))

#         # page.pause()
#         page.locator("iframe[name=\"appFrame\"]").content_frame.get_by_text("Your browser does not support").content_frame.locator("#actionMenu").click(timeout = 1000000)
        
        
#         with page.expect_download() as download_info:
#             page.locator("iframe[name=\"appFrame\"]").content_frame.get_by_text("Your browser does not support").content_frame.get_by_text("Baixar CSV").click(timeout = 1000000)
#         q.put(("status", "Inicializando download do modelo {key}"))
#         download = download_info.value

#         q.put(("status", "Download realizado com sucesso do modelo {key}"))
#         file_path = f"Dados/{key}.csv"
#         os.makedirs("Dados", exist_ok=True)
        
#         if os.path.exists(file_path):
#             os.remove(file_path)
        
#         download.save_as(file_path)
#         q.put(("status", f"Relatório {key} salvo como: {file_path}"))

#     page.pause()






def download_por_modelo(page, url_oss, q, username, password, Modelos):
    page.goto(url_oss, timeout=600000)
    q.put(("status", "Realizando login..."))
    q.put(("progress", 15))

    page.locator('[name="USER_NAME"]').fill(username)
    page.locator('[name="PASSWORD"]').fill(password)
    page.locator(".signin").click()
    q.put(("status", "Login realizado com sucesso!"))

    for key, value in Modelos.items():
        q.put(("status", f"Processando modelo {key}"))

        # Skip known failing model
        if key == '611':
            continue

        page.locator(".shellInstance").click()
        print(key, value)

        page.locator("#sequencer_ui_instances").select_option(value)
        page.get_by_role("link", name="Editor de programação").click(timeout=500000)

        q.put(("status", f"Aguardando carregar página de relatório do modelo {key}"))

        frame = page.locator("iframe[name=\"appFrame\"]").content_frame
        frame.get_by_text("Your browser does not support").content_frame.locator("#actionMenu").click(timeout=1000000)

        with page.expect_download() as download_info:
            frame.get_by_text("Your browser does not support").content_frame.get_by_text("Baixar CSV").click(timeout=1000000)
        q.put(("status", f"Inicializando download do modelo {key}"))
        download = download_info.value

        os.makedirs("Dados", exist_ok=True)
        csv_path = f"Dados/{key}.csv"
        xlsx_path = f"Dados/{key}.xlsx"

        # Clean old files
        for path in [csv_path, xlsx_path]:
            if os.path.exists(path):
                os.remove(path)

        download.save_as(csv_path)
        q.put(("status", f"Relatório {key} salvo como: {csv_path}"))

        # ✅ Convert to Excel only for rows where order_type == 'PREV'
        try:
            df = pd.read_csv(csv_path, low_memory=False)
            
            if "order_type" in df.columns:
                df_prev = df[df["order_type"] == "PRE"]

                if not df_prev.empty:
                    df_prev.to_excel(xlsx_path, index=False, engine="xlsxwriter")
                    q.put(("status", f"Arquivo Excel criado com {len(df_prev)} linhas: {xlsx_path}"))
                else:
                    q.put(("status", f"Nenhuma linha com order_type='PREV' no modelo {key}, conversão ignorada."))
            else:
                q.put(("status", f"Coluna 'order_type' não encontrada no modelo {key}, conversão ignorada."))
        except Exception as e:
            q.put(("status", f"Erro ao converter {key} para Excel: {e}"))


# def check_file(q) :
#     xlsx_path = "Dados/341.xlsx"
   
#     try:
#         df = pd.read_csv("Dados/341.csv", low_memory=False)
#         print(df["order_type"].unique()[:2000])
        
#         if "order_type" in df.columns:
#             df_prev = df[df["order_type"] == "PRE"]

#             if not df_prev.empty:
#                 df_prev.to_excel(xlsx_path, index=False, engine="xlsxwriter")
#                 q.put(("status", f"Arquivo Excel criado com {len(df_prev)} linhas: {xlsx_path}"))
#             else:
#                 q.put(("status", f"Nenhuma linha com order_type='PREV' no modelo , conversão ignorada."))
#         else:
#             q.put(("status", f"Coluna 'order_type' não encontrada no modelo , conversão ignorada."))
#     except Exception as e:
#         q.put(("status", f"Erro ao converter para Excel: {e}"))