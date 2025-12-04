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
from Tasks import download_por_modelo,Atualizar_Links_Pivort_tables_Single_Model
from Tasks import download_A14  # Kept for future use


warnings.filterwarnings("ignore", category=UserWarning)

# --- HELPER FUNCTIONS (UNCHANGED) ---


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


def update_gui(queue_instance, status_label, progress_bar, log_text, root=None, process_button=None):
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
                if process_button:
                    process_button.config(state="normal")
                return # Stop the polling loop
    except queue.Empty:
        pass
    
    # Continue polling
    if root:
        root.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text, root, process_button))
    else:
        status_label.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text))


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
       
        # Atualizar_Links_Pivort_tables_Single_Model("265",q)
        # return
    
        # Data validation check
        if not isinstance(modelos_to_process, dict):
            error_msg = "ERRO: O arquivo Modelos.json não é um dicionário válido."
            q.put(("status", error_msg))
            raise TypeError(error_msg)
            
        chromium_path = get_playwright_browser_path()
        
        q.put(("status", "Iniciando download dos arquivos..."))

        # # Create and start the manager thread for downloading models.
        model_thread = threading.Thread(
            target=download_por_modelo, 
            args=(url_oss, q, username, password, modelos_to_process, chromium_path)
        )
        
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
        self.root.title("Ferramenta de Automação e Processamento RPA")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        
        dhl_red = "#FF0000"
        dhl_yellow = "#FFCC00"
        stellantis_blue = "#003DA5"
        stellantis_orange = "#FF6600"
        
        # Set modern color scheme with DHL/STELLANTIS theme
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure button style with STELLANTIS blue
        style.configure('TButton', background=stellantis_blue, foreground="white", relief="flat", padding=6)
        style.map('TButton', background=[('active', stellantis_orange)])
        
        # Configure progressbar with gradient effect using STELLANTIS colors
        style.configure('TProgressbar', background=stellantis_blue, troughcolor='#E8E8E8', bordercolor='#CCCCCC', lightcolor=stellantis_orange, darkcolor=stellantis_blue)
        
        # Configure labels with theme colors
        style.configure('Title.TLabel', font=("Segoe UI", 16, "bold"), foreground=stellantis_blue)
        
        self.queue = queue.Queue()

        # --- Main container ---
        container = tk.Frame(root, bg="white")
        container.pack(fill=tk.BOTH, expand=True)

        # --- Header with DHL/STELLANTIS accent ---
        header_frame = tk.Frame(container, bg=stellantis_blue, height=80)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Title section with colored background
        title_label = tk.Label(header_frame, text="🤖 Automação e Processamento Avançados", font=("Segoe UI", 16, "bold"), fg="white", bg=stellantis_blue)
        title_label.pack(anchor="w", padx=15, pady=(10, 2))
        
        subtitle_label = tk.Label(header_frame, text="Processamento Inteligente de Processos Manuais", font=("Segoe UI", 9), fg=dhl_yellow, bg=stellantis_blue)
        subtitle_label.pack(anchor="w", padx=15, pady=(0, 10))

        # --- Main content frame ---
        main_frame = ttk.Frame(container, padding="13")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status section
        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Segoe UI", 11), foreground=stellantis_blue)
        self.status_label.pack(pady=(2, 3), padx=1, fill=tk.X)

        # Progress bar with accent color
        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        # Button section with modern styling
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=4, fill=tk.X)
        
        self.process_button = ttk.Button(button_frame, text="▶ Processar", command=self.start_processing_thread, style='TButton')
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        # Log section with accent
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Atividades", padding="11")
        log_frame.pack(pady=0, padx=2, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=12, font=("Consolas",11), bg="#F5F5F5", fg="#333333")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Footer section with DHL/STELLANTIS branding
        footer_frame = tk.Frame(container, bg=stellantis_blue, height=34)
        footer_frame.pack(fill=tk.X, padx=0, pady=0, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Left side - DHL -> STELLANTIS
        left_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        left_footer.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=10)
        
        # DHL Logo/Text (DHL Yellow)
        dhl_label = tk.Label(left_footer, text="🚚 DHL", font=("Segoe UI", 11, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        dhl_label.pack(side=tk.LEFT, padx=1)
        
        arrow_label = tk.Label(left_footer, text="→", font=("Segoe UI", 12, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        arrow_label.pack(side=tk.LEFT, padx=3)
        
        # STELLANTIS Logo/Text (STELLANTIS Orange accent)
        stellantis_label = tk.Label(left_footer, text="STELLANTIS 🏢", font=("Segoe UI", 11, "bold"), fg=stellantis_orange, bg=stellantis_blue)
        stellantis_label.pack(side=tk.LEFT, padx=3)
        
        # Right side - Developer credit
        right_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        right_footer.pack(side=tk.RIGHT, padx=15, pady=10)
        
        footer_label = tk.Label(right_footer, text="Desenvolvido por: Vincent Pernarh", font=("Segoe UI", 9), fg="white", bg=stellantis_blue)
        footer_label.pack(anchor="e")

    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo...")
        
        self.thread = threading.Thread(target=main_process, args=(self.queue,))
        self.thread.daemon = True
        self.thread.start()
        
        # Start checking the queue for updates
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text, self.root, self.process_button)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()