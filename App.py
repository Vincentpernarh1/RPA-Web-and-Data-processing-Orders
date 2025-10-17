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

def get_playwright_browser_path():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        chromium_path = os.path.join(base_path, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe")
    else:
        base_path = r"C:\Users\perna\AppData\Local"

        # Join the rest of the Playwright folder path
        chromium_path = os.path.join(
            base_path,
            "ms-playwright",
            "chromium-1187",
            "chrome-win",
            "chrome.exe"
        )
   
    if chromium_path and not os.path.exists(chromium_path):
        raise FileNotFoundError(f"Chromium executable not found at {chromium_path}")

    return chromium_path


# --- GUI UPDATE FUNCTION ---
def update_gui(queue_instance, status_label, progress_bar, log_text):
    """Checks the queue for messages from the worker thread and updates the GUI."""
    try:
        while True:
            message_type, value = queue_instance.get_nowait()
            if message_type == "status":
                status_label.config(text=value)
                log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {value}\n")
                log_text.see(tk.END)
            elif message_type == "progress":
                progress_bar['value'] = value
            elif message_type == "done":
                status_label.config(text="Processo Concluído!")
                progress_bar['value'] = 100
                return # Stop checking
    except queue.Empty:
        pass
    status_label.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text))


def load_credentials():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    cred_path = os.path.join(base_path, "credencial.json")

    if not os.path.exists(cred_path):
        raise FileNotFoundError(f"Credencial.json not found in: {cred_path}")

    with open(cred_path, "r", encoding="utf-8") as f:
        return json.load(f)
    

def load_modelos():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    Model_path = os.path.join(base_path, "Modelos.json")

    if not os.path.exists(Model_path):
        raise FileNotFoundError(f"Credencial.json not found in: {Model_path}")

    with open(Model_path, "r", encoding="utf-8") as f:
        return json.load(f)
    


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



def download_por_modelo(page,url_oss,q,username,password,Modelos) :
    page.goto(url_oss, timeout= 600000 );
    # 3. Login
    q.put(("status", "Realizando login..."))
    q.put(("progress", 15))

    page.locator('[name="USER_NAME"]').fill(username)
    page.locator('[name="PASSWORD"]').fill(password)
    page.locator(".signin").click()
    q.put(("status", "Login realizado com sucesso!"))
    

    for key, value in Modelos.items(): 
        q.put(("status", "Processando modelo {key}"))
        if key == '611' :    #this line is for us to jump the 611 models and its report is failling causing crash, i could just add try andf catch to prevent crash.
            continue

        page.locator(".shellInstance").click()
        print(key,value)

        page.locator("#sequencer_ui_instances").select_option(value)
        page.get_by_role("link", name="Editor de programação").click(timeout=500000)

        q.put(("status", "Aguardando carregar pagina de relatório do modelo {key}"))

        page.pause()
        page.locator("iframe[name=\"appFrame\"]").content_frame.get_by_text("Your browser does not support").content_frame.locator("#actionMenu").click(timeout = 1000000)
        
        
        with page.expect_download() as download_info:
            page.locator("iframe[name=\"appFrame\"]").content_frame.get_by_text("Your browser does not support").content_frame.get_by_text("Baixar CSV").click(timeout = 1000000)
        q.put(("status", "Inicializando download do modelo {key}"))
        download = download_info.value

        q.put(("status", "Download realizado com sucesso do modelo {key}"))
        file_path = f"Dados/{key}.csv"
        os.makedirs("Dados", exist_ok=True)
        
        if os.path.exists(file_path):
            os.remove(file_path)
        
        download.save_as(file_path)
        q.put(("status", f"Relatório {key} salvo como: {file_path}"))

    page.pause()


    
def run_automation(playwright: Playwright, q: queue.Queue):
    ecr_path, odm_path = None, None
    try:
        # 1. Load Credentials
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 5))
        credentials = load_credentials()
        url_order, username, password ,url_oss= credentials['url_order'], credentials['user'], credentials['password'],credentials['url_oss']

        # 2. Launch Browser
        q.put(("status", "Iniciando navegador..."))

        chromium_path = get_playwright_browser_path()
        
        if chromium_path:
            # .exe → use bundled Chromium
            browser = playwright.chromium.launch(
                headless=False,
                executable_path=chromium_path,
                args=["--start-maximized"]
            )
        else:
            # .py → use default Playwright Chromium
            browser = playwright.chromium.launch(
                headless=True,
                args=["--start-maximized"]
            )
                    
        # context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        # download_A14(page,url_order,q,username,password)
        download_por_modelo(page,url_oss,q,username,password,Modelos=load_modelos())




       

    except FileNotFoundError:
        q.put(("status", "Erro: 'Credencial.json' não encontrado."))
    except KeyError:
        q.put(("status", "Erro: JSON de credenciais inválido."))
    except TimeoutError:
        q.put(("status", "Erro de Timeout: Verifique os seletores ou a conexão."))
        page.screenshot(path="login_error.png")
    except Exception as e:
        q.put(("status", f"Ocorreu um erro inesperado: {e}"))
    finally:
        # 5. Clean Up and next step
        q.put(("status", "Fechando navegador..."))
        if 'context' in locals(): context.close()
        if 'browser' in locals(): browser.close()
        
        

        q.put(("done", True))



def main_process(q: queue.Queue):
    with sync_playwright() as playwright:
        run_automation(playwright, q)

# --- TKINTER APP SETUP ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Automação e Processamento")
        self.root.geometry("600x400")

        self.queue = queue.Queue()

        # --- Widgets ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Helvetica", 12))
        self.status_label.pack(pady=5, padx=5, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        self.process_button = ttk.Button(main_frame, text="Processar", command=self.start_processing_thread)
        self.process_button.pack(pady=10)
        
        log_frame = ttk.LabelFrame(main_frame, text="Log de Atividades", padding="10")
        log_frame.pack(pady=10, padx=5, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=70, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo...")
        
        self.thread = threading.Thread(target=main_process, args=(self.queue,))
        self.thread.daemon = True
        self.thread.start()
        
        # Start checking the queue for updates
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
