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
from playwright.sync_api import Page, Browser

warnings.filterwarnings("ignore", category=UserWarning)
import sys


def Process_A14_options(file_path, q):
    q.put(("status", "🔄 Inicializando o processamento dos arquivos..."))

    # Step 1: Load data (This part is unchanged)
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".xlsx", ".xlsm"]:
            df = pd.read_excel(file_path, engine="openpyxl", dtype=object)
        elif ext == ".xls":
            df = pd.read_excel(file_path, engine="xlrd", dtype=object)
        elif ext == ".xlsb":
            df = pd.read_excel(file_path, engine="pyxlsb", dtype=object)
        elif ext == ".csv":
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                sample = f.read(4096)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
                    delimiter = dialect.delimiter
                except csv.Error:
                    delimiter = ";" if ";" in sample else ","
            try:
                df = pd.read_csv(file_path, delimiter=delimiter, encoding="utf-8", engine="python", dtype=object)
            except Exception:
                df = pd.read_csv(file_path, delimiter=delimiter, encoding="latin-1", engine="python", dtype=object)
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
        # Preserve exact string representation - no conversion whatsoever
        pack_value = row[pack_col_name]
        if pd.notna(pack_value):
            # Convert to string but preserve the original representation
            pack = str(pack_value).strip() if str(pack_value).strip().lower() != 'nan' else ""
        else:
            pack = ""
        
        # Step 4: Get and join all CONTEÚDO values from the other optional columns
        conteudo_values = []
        for col in conteudo_cols:
            value = row[col]
            if pd.notna(value):
                str_value = str(value).strip()
                # Exclude NaN strings
                if str_value and str_value.lower() != 'nan':
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
                app = xw.App(visible=True, add_book=False)
                app.display_alerts = False
                app.screen_updating = True

                wb = app.books.open(file_full_path, update_links=False)
                if 'A14' in [s.name for s in wb.sheets]:
                    ws = wb.sheets['A14']
                else:
                    ws = wb.sheets.add('A14')
                ws.clear_contents()
                
                # Write headers
                ws.range('A1').value = ['PACK', 'CONTEÚDO']
                
                # Format columns as text to preserve string format
                ws.range('A:B').number_format = '@'
                
                # Write data
                ws.range('A2').value = df_result.values.tolist()
                
                ws.autofit()
                wb.save()
                wb.close()
                # app.quit()
                q.put(("status", f"✅ Planilha 'A14' atualizada em {filename}"))
            except Exception as e:
                q.put(("status", f"❌ Falha ao processar {filename}: {e}"))

    q.put(("status", "🎉 Processamento concluído com sucesso."))


def download_A14(url_order,q,username,password,chromium_path) :
    with sync_playwright() as p:
        browser = None
        try :
            browser = p.chromium.launch(
                    headless=False, # Always headless for worker threads
                    executable_path=chromium_path,
                    # args=["--start-maximized"]
                )
            context = browser.new_context(viewport={'width': 1920, 'height': 1080})
            page = context.new_page()
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
        except Exception as e:
            q.put(("status", f"ERROR ao processar o arquivo A14: {e}"))
        finally:
            pass




def download_por_modelo(url_oss, q, username, password, Modelos, chromium_path):
    """This function is the 'thread manager'. It creates and starts a thread for each model."""
    threads = []
    
    for key, value in Modelos.items():
        if key == '611':
            q.put(("status", "Skipping known failing model 611."))
            continue
            
        thread = threading.Thread(
            target=process_single_model,
            # ARGS TUPLE IS NOW CORRECT: No browser, but chromium_path is added at the end
            args=(url_oss, q, username, password, key, value, chromium_path)
        )
        threads.append(thread)
        thread.start()
        q.put(("status", f"Manager: Started thread for model {key}."))

    # Wait for all model-processing threads to complete
    for thread in threads:
        thread.join()

    q.put(("status", "Manager: All model processing threads have finished."))


def process_single_model(url_oss: str, q: queue.Queue, username: str, password: str, key: str, value: str, chromium_path: str):
    """
    Processes a single model.
    This function is now COMPLETELY independent and thread-safe.
    """
    # Each thread now creates its own Playwright instance and browser.
    with sync_playwright() as p:
        browser = None
        try:
            browser = p.chromium.launch(
                headless=False, # Always headless for worker threads
                executable_path=chromium_path,
                # args=["--start-maximized"]
            )
            context = browser.new_context(viewport={'width': 1920, 'height': 1080})
            page = context.new_page()

            q.put(("status", f"Model {key}: Thread starting..."))
            page.goto(url_oss, timeout=60000)

            page.locator('[name="USER_NAME"]').fill(username)
            page.locator('[name="PASSWORD"]').fill(password)
            page.locator(".signin").click()

            page.locator(".shellInstance").click()
            page.locator("#sequencer_ui_instances").select_option(value)
            page.get_by_role("link", name="Editor de programação").click(timeout=120000)

            q.put(("status", f"Model {key}: Waiting for report page..."))

            frame = page.locator("iframe[name=\"appFrame\"]").content_frame
            inner_frame = frame.get_by_text("Your browser does not support").content_frame
            
            inner_frame.locator("#actionMenu").click(timeout=120000)

            with page.expect_download(timeout=120000) as download_info:
                inner_frame.get_by_text("Baixar CSV").click()
            
            download = download_info.value
            q.put(("status", f"Model {key}: Download initiated."))

            os.makedirs("Dados", exist_ok=True)
            csv_path = os.path.join("Dados", f"{key}.csv")
            xlsx_path = os.path.join("Dados", f"{key}.xlsx")

            if os.path.exists(csv_path): os.remove(csv_path)
            if os.path.exists(xlsx_path): os.remove(xlsx_path)

            # download.save_as(csv_path)
            temp_csv_path =download.path()
            q.put(("status", f"Model {key}: Saved to {csv_path}"))
            
            df = pd.read_csv(temp_csv_path, low_memory=False)
            if "order_type" in df.columns and not df[df["order_type"] == "PRE"].empty:
                df_prev =df[df["order_type"] == "PRE"]
                df_prev.to_excel(xlsx_path, index=False, engine="xlsxwriter")

                q.put(("status", f"Model {key}: Excel created at {xlsx_path}"))

                q.put(("status", f"Atualizando planilha Base para o Modelo {key}"))
                Atualizar_Base_previsao(df_prev, key, q)
                q.put(("status", f"Planilha Base do Modelo {key}: Atualizado com sucesso "))
            
            q.put(("status", f"Model {key}: Finished successfully."))

        except Exception as e:
            q.put(("status", f"ERROR in model thread {key}: {e}"))
        finally:
            # The 'with sync_playwright()' block handles all cleanup automatically.
            # No need to manually close the browser here.
            pass

# In Tasks.py, add this new function
import os
import sys
import xlwings as xw
import pandas as pd

def Atualizar_Base_previsao(df_to_paste, model_key, q):
    """
    Finds the correct 'BASE' file, clears it, pastes the data (without the header),
    copies formatting from the second row, and autofills formulas.
    """
    q.put(("status", f"Model {model_key}: Procurando arquivo BASE para atualizar..."))
    
    try:
        # Step 1: Find the target file in the 'Bases' subfolder
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        bases_folder = os.path.join(base_path, "Bases")

        if not os.path.exists(bases_folder):
            q.put(("status", f"Model {model_key}: ERRO - Pasta 'Bases' não encontrada."))
            return

        target_file_path = None
        target_filename = None
        for filename in os.listdir(bases_folder):
            if (filename.upper().startswith('BASE') and 
                model_key in filename and 
                not filename.startswith("~")):
                
                target_file_path = os.path.join(bases_folder, filename)
                target_filename = filename
                break

        if not target_file_path:
            q.put(("status", f"Model {model_key}: ERRO - Nenhum arquivo BASE encontrado para este modelo."))
            return
            
        q.put(("status", f"Model {model_key}: Arquivo encontrado: {target_filename}"))

        # Use a new app instance for each thread to ensure stability
        with xw.App(visible=False) as app:
            wb = app.books.open(target_file_path)
            try:
                ws = wb.sheets['ARQUIVO PREVISÕES']
                
                # --- NEW LOGIC START ---

                # Prepare the data: Remove only the first row (the header).
                df_data_only = df_to_paste.iloc[1:]

                # Always clear the old data content first.
                ws.range('B2:Y1048576').clear_contents()

                # Proceed only if there is data left after removing the header.
                if not df_data_only.empty:
                    # 1. PASTE NEW DATA
                    start_cell = ws.range('B2')
                    start_cell.options(header=False, index=False).value = df_data_only
                    q.put(("status", f"Model {model_key}: Dados colados com sucesso em '{target_filename}'."))

                    # Get the full range of the data we just pasted.
                    pasted_range = start_cell.expand('table')

                    # 2. APPLY FORMATTING
                    # If more than one row was pasted, copy the format from the first new row to the rest.
                    if pasted_range.rows.count > 1:
                        # The source of the format is the entire first row of our pasted data (Row 2).
                        format_source_range = pasted_range.rows[0]

                        # The destination is all other pasted rows (Row 3 downwards).
                        format_destination_range = ws.range(pasted_range.rows[1], pasted_range.rows[-1])
                        
                        # Copy and paste_special to apply only the formats.
                        format_source_range.copy()
                        format_destination_range.paste(paste='formats')
                        app.api.CutCopyMode = False # Clear the clipboard to prevent Excel warnings
                        q.put(("status", f"Model {model_key}: Formato aplicado em {pasted_range.rows.count - 1} linhas."))

                    # 3. AUTOFILL FORMULAS
                    last_row = pasted_range.last_cell.row
                    
                    # !!! ACTION REQUIRED !!!
                    # Change the column letters below to your actual formula columns.
                    formula_columns = ['Z', 'AA', 'AB'] # <-- EDIT THESE COLUMNS

                    for col in formula_columns:
                        formula_source_range = ws.range(f'{col}2')
                        formula_fill_range = ws.range(f'{col}2:{col}{last_row}')
                        
                        # Use autofill to drag down the formulas from row 2.
                        formula_source_range.autofill(destination=formula_fill_range)

                    q.put(("status", f"Model {model_key}: Fórmulas preenchidas até a linha {last_row}."))

                else:
                    # If there's no data after removing the header, log a warning.
                    q.put(("status", f"Model {model_key}: AVISO - Sem dados para colar. Planilha '{target_filename}' foi limpa."))
                
                # --- NEW LOGIC END ---

                wb.save()
            except Exception as sheet_error:
                q.put(("status", f"Model {model_key}: ERRO ao processar a planilha em '{target_filename}': {sheet_error}"))
            finally:
                wb.close()

    except Exception as e:
        q.put(("status", f"Model {model_key}: ERRO FATAL em Atualizar_Base_previsao: {e}"))


    """
    Finds the correct 'BASE' file, clears it, and pastes the processed 
    DataFrame if it contains enough data after slicing.
    """
    q.put(("status", f"Model {model_key}: Procurando arquivo BASE para atualizar..."))
    
    try:
        # Step 1: Find the target file in the 'Bases' subfolder
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        bases_folder = os.path.join(base_path, "Bases")

        if not os.path.exists(bases_folder):
            q.put(("status", f"Model {model_key}: ERRO - Pasta 'Bases' não encontrada."))
            return

        target_file_path = None
        target_filename = None
        for filename in os.listdir(bases_folder):
            if (filename.upper().startswith('BASE') and 
                model_key in filename and 
                not filename.startswith("~")):
                
                target_file_path = os.path.join(bases_folder, filename)
                target_filename = filename
                break

        if not target_file_path:
            q.put(("status", f"Model {model_key}: ERRO - Nenhum arquivo BASE encontrado para este modelo."))
            return
            
        q.put(("status", f"Model {model_key}: Arquivo encontrado: {target_filename}"))

        with xw.App(visible=False) as app:
            wb = app.books.open(target_file_path)
            try:
                ws = wb.sheets['ARQUIVO PREVISÕES']
                
                # --- MODIFIED SECTION START ---

                # Always clear the old data first.
                ws.range('B2').expand().clear_contents()

                # Check if the DataFrame has more than 1 row AND more than 1 column.
                if df_to_paste.shape[0] > 1 and df_to_paste.shape[1] > 1:
                    # If so, slice to remove the first row and first column.
                    df_final_to_paste = df_to_paste.iloc[1:, 1:]
                    
                    # Paste the processed data.
                    ws.range('B2').options(header=False, index=False).value = df_final_to_paste
                    q.put(("status", f"Model {model_key}: Base '{target_filename}' atualizada com sucesso."))
                
                else:
                    # If the DataFrame is too small, there's nothing to paste after slicing.
                    q.put(("status", f"Model {model_key}: AVISO - Dados de origem insuficientes. Planilha '{target_filename}' foi limpa."))

                # --- MODIFIED SECTION END ---
                
                wb.save()
            except Exception as sheet_error:
                q.put(("status", f"Model {model_key}: ERRO ao processar a planilha em '{target_filename}': {sheet_error}"))
            finally:
                wb.close()

    except Exception as e:
        q.put(("status", f"Model {model_key}: ERRO FATAL em Atualizar_Base_previsao: {e}"))
    """
    Finds the correct 'BASE' file for a given model and pastes the provided DataFrame into it.
    """
    q.put(("status", f"Model {model_key}: Procurando arquivo BASE para atualizar..."))
    
    try:
        # Step 1: Find the target file in the 'Bases' subfolder
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        bases_folder = os.path.join(base_path, "Bases")

        if not os.path.exists(bases_folder):
            q.put(("status", f"Model {model_key}: ERRO - Pasta 'Bases' não encontrada."))
            return

        target_file_path = None
        target_filename = None
        for filename in os.listdir(bases_folder):
            # Find a file that starts with 'BASE' and contains the model number
            if (filename.upper().startswith('BASE') and 
                model_key in filename and 
                not filename.startswith("~")):
                
                target_file_path = os.path.join(bases_folder, filename)
                target_filename = filename
                break # Stop after finding the first match

        if not target_file_path:
            q.put(("status", f"Model {model_key}: ERRO - Nenhum arquivo BASE encontrado para este modelo."))
            return
            
        q.put(("status", f"Model {model_key}: Arquivo encontrado: {target_filename}"))

        # Step 2: Open the file with xlwings and paste the data
        # Use a new app instance for each thread to ensure stability
        with xw.App(visible=False) as app:
            wb = app.books.open(target_file_path)
            try:
                ws = wb.sheets['ARQUIVO PREVISÕES']
                
                # Clear old data and paste new data starting at B2
                # This automatically includes the header
                ws.range('B2:Y10000000000').expand().clear_contents()
                ws.range('B2').value = df_to_paste[1:,]
                
                wb.save()
                q.put(("status", f"Model {model_key}: Base '{target_filename}' atualizada com sucesso."))
            except Exception as sheet_error:
                q.put(("status", f"Model {model_key}: ERRO ao processar a planilha em '{target_filename}': {sheet_error}"))
            finally:
                wb.close()

    except Exception as e:
        q.put(("status", f"Model {model_key}: ERRO FATAL em Atualizar_Base_previsao: {e}"))