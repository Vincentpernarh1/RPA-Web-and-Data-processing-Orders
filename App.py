import json
import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import sys
import warnings

# --- Import your custom task functions ---
# Make sure your Tasks.py file is in the same folder
from Tasks import download_por_modelo
from Tasks import download_A14  # Kept for future use


warnings.filterwarnings("ignore", category=UserWarning)

# --- HELPER FUNCTIONS (UNCHANGED) ---

def get_playwright_browser_path():
    """Determines the path to the Playwright Chromium executable."""
    if getattr(sys, 'frozen', False):
        # Path when running as a bundled executable (e.g., PyInstaller)
        base_path = sys._MEIPASS
    else:
        # Path when running as a .py script
        base_path = os.path.expanduser(r"~\AppData\Local")

    chromium_path = os.path.join(
        base_path, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe"
    )
    
    if not os.path.exists(chromium_path):
        # Fallback if the specific version is not found, letting Playwright decide.
        # This makes the script more robust to Playwright updates.
        return None
    
    return chromium_path


def update_gui(queue_instance, status_label, progress_bar, log_text, root):
    """Checks the queue for messages and updates the GUI."""
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
                # Re-enable the button when done
                root.nametowidget("main_frame.process_button").config(state="normal")
                return # Stop the polling loop
    except queue.Empty:
        pass
    
    # Continue polling
    root.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text, root))


def load_credentials():
    """Loads credencial.json from the script's directory."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    cred_path = os.path.join(base_path, "credencial.json")
    if not os.path.exists(cred_path):
        raise FileNotFoundError(f"credencial.json not found in: {cred_path}")
    with open(cred_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_modelos():
    """Loads Modelos.json from the script's directory."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    model_path = os.path.join(base_path, "Modelos.json")
    if not os.path.exists(model_path):
        raise FileNotFoundError(f"Modelos.json not found in: {model_path}")
    with open(model_path, "r", encoding="utf-8") as f:
        return json.load(f)

# --- CORRECTED MAIN AUTOMATION LOGIC ---

def main_process(q: queue.Queue):
    """
    This is the main function targeted by the GUI thread.
    It orchestrates the automation tasks in a thread-safe way.
    """
    try:
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 5))
        credentials = load_credentials()
        url_order, username, password, url_oss = credentials['url_order'], credentials['user'], credentials['password'], credentials['url_oss']
        
        modelos_to_process = load_modelos()
       
        
        # Data validation check
        if not isinstance(modelos_to_process, dict):
            error_msg = "ERRO: O arquivo Modelos.json não é um dicionário válido."
            q.put(("status", error_msg))
            raise TypeError(error_msg)
            
        chromium_path = get_playwright_browser_path()
        
        q.put(("status", "Iniciando download dos arquivos..."))

        # Create and start the manager thread for downloading models.
        model_thread = threading.Thread(
            target=download_por_modelo, 
            args=(url_oss, q, username, password, modelos_to_process, chromium_path)
        )
        

        # For example:
        a14_thread = threading.Thread(target=download_A14,  
                args=(url_order, q, username, password, chromium_path)
                )
        


        # Inicializando os threads
        a14_thread.start()
        model_thread.start()
        
        # Wait for all threads to finish
        model_thread.join()
        a14_thread.join()

        q.put(("status", "Processo de automação finalizado."))

    except (FileNotFoundError, KeyError, TypeError) as e:
        q.put(("status", f"ERRO CRÍTICO: {e}"))
    except Exception as e:
        q.put(("status", f"Ocorreu um erro inesperado: {e}"))
    finally:
        q.put(("done", True))

# --- TKINTER APP SETUP (UNCHANGED) ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Automação e Processamento")
        self.root.geometry("600x400")
        self.queue = queue.Queue()

        main_frame = ttk.Frame(root, padding="10", name="main_frame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Helvetica", 12))
        self.status_label.pack(pady=5, padx=5, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        self.process_button = ttk.Button(main_frame, text="Processar", command=self.start_processing_thread, name="process_button")
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
        
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text, self.root)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()