import os
import re
import sys
import threading
import time
import json
from pathlib import Path
from datetime import timedelta
import shutil
import subprocess
import platform
import unicodedata

try:
    import pandas as pd
except ImportError:
    os.system("pip install pandas openpyxl xlrd")
    import pandas as pd

try:
    import PyPDF2
except ImportError:
    os.system("pip install PyPDF2")
    import PyPDF2

try:
    import pdfplumber
except ImportError:
    os.system("pip install pdfplumber")
    import pdfplumber

try:
    import customtkinter as ctk
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from PIL import Image, ImageTk
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
except ImportError:
    try:
        import customtkinter as ctk
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        Image = None
        ImageTk = None
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
    except ImportError:
        print("Erro: customtkinter não instalado. Execute: pip install customtkinter")
        sys.exit(1)


# ==================== RESOURCE PATH HELPER ====================

def resource_path(relative_path):
    """
    Obtém o caminho absoluto para recursos, funciona tanto em desenvolvimento
    quanto quando empacotado pelo PyInstaller.
    
    Quando o PyInstaller cria um executável, ele descompacta os recursos em uma
    pasta temporária e armazena o caminho em sys._MEIPASS.
    """
    try:
        # PyInstaller cria uma pasta temp e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Em desenvolvimento, usa o diretório atual
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


# ==================== GOOGLE DRIVE UPLOAD DIALOGS ====================

class DriveUploadDialog:
    """Janela de revisão e configuração antes do upload"""

    def __init__(self, parent, app, source_folder, file_summary):
        self.parent = parent
        self.app = app
        self.source_folder = source_folder
        self.file_summary = file_summary
        self.result = None

        self.window = ctk.CTkToplevel(parent)
        self.window.title("Enviar para Google Drive")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.geometry("1200x800")
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - 600
        y = (self.window.winfo_screenheight() // 2) - 400
        self.window.geometry(f"+{x}+{y}")

        self.setup_ui()

    def setup_ui(self):
        ACCENT = "#00A8CC"
        main = ctk.CTkFrame(self.window, fg_color="transparent")
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Cabeçalho
        ctk.CTkLabel(main, text="📤 Enviar Comprovantes para Google Drive",
                     font=("Segoe UI", 16, "bold"), text_color=ACCENT).pack(pady=(0, 5))
        ctk.CTkLabel(main, text="Revise os arquivos e selecione o destino antes de enviar",
                     font=("Segoe UI", 10)).pack(pady=(0, 15))

        # Resumo
        summary_frame = ctk.CTkFrame(main, corner_radius=8)
        summary_frame.pack(fill=tk.X, pady=(0, 12))
        ctk.CTkLabel(summary_frame, text="📊 Resumo",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 5))
        info_text = (f"📁 Pasta origem: {os.path.basename(self.source_folder)}\n"
                     f"📄 Total de arquivos: {self.file_summary['total_files']}\n"
                     f"📂 Centros de custo: {self.file_summary['total_folders']}\n"
                     f"💾 Tamanho total: {self.app.format_size(self.file_summary['total_size'])}")
        ctk.CTkLabel(summary_frame, text=info_text, justify=tk.LEFT).pack(anchor=tk.W, padx=15, pady=(0, 10))

        # Lista de pastas (Treeview - sem equivalente no CTk)
        list_frame = ctk.CTkFrame(main, corner_radius=8)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))
        ctk.CTkLabel(list_frame, text="📋 Arquivos por Centro de Custo",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 5))

        tree_container = ctk.CTkFrame(list_frame, fg_color="transparent")
        tree_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        tree = ttk.Treeview(tree_container, columns=('files', 'size'), show='tree headings', selectmode='none')
        tree.heading('#0', text='Centro de Custo')
        tree.heading('files', text='Arquivos')
        tree.heading('size', text='Tamanho')
        tree.column('#0', width=400)
        tree.column('files', width=100, anchor=tk.CENTER)
        tree.column('size', width=150, anchor=tk.CENTER)
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for ccusto, data in sorted(self.file_summary['folders'].items()):
            tree.insert('', 'end', text=f"✓ {ccusto}",
                        values=(data['count'], self.app.format_size(data['size'])))

        # Destino
        dest_frame = ctk.CTkFrame(main, corner_radius=8)
        dest_frame.pack(fill=tk.X, pady=(0, 12))
        ctk.CTkLabel(dest_frame, text="🎯 Destino no Google Drive",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 5))

        self.drive_path = tk.StringVar()
        detected = self.app.detect_google_drive_folder()
        if detected:
            self.drive_path.set(detected)
            ctk.CTkLabel(dest_frame, text="✓ Google Drive detectado automaticamente",
                         text_color="#4CAF50", font=("Segoe UI", 9)).pack(anchor=tk.W, padx=15, pady=(0, 5))

        path_frame = ctk.CTkFrame(dest_frame, fg_color="transparent")
        path_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        ctk.CTkLabel(path_frame, text="Pasta:").pack(side=tk.LEFT, padx=(0, 10))
        ctk.CTkEntry(path_frame, textvariable=self.drive_path, font=("Segoe UI", 10)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ctk.CTkButton(path_frame, text="📁 Procurar...", command=self.select_drive_folder,
                      width=120, fg_color=ACCENT, hover_color="#0088AA").pack(side=tk.LEFT)

        # Opções
        options_frame = ctk.CTkFrame(main, corner_radius=8)
        options_frame.pack(fill=tk.X, pady=(0, 12))
        ctk.CTkLabel(options_frame, text="⚙️ Opções",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 5))

        self.keep_local = tk.BooleanVar(value=True)
        self.create_backup = tk.BooleanVar(value=False)
        self.open_after = tk.BooleanVar(value=True)

        opts_inner = ctk.CTkFrame(options_frame, fg_color="transparent")
        opts_inner.pack(anchor=tk.W, padx=15, pady=(0, 10))
        ctk.CTkCheckBox(opts_inner, text="Manter cópia local após upload", variable=self.keep_local).pack(anchor=tk.W, pady=2)
        ctk.CTkCheckBox(opts_inner, text="Criar backup antes de enviar (.zip)", variable=self.create_backup).pack(anchor=tk.W, pady=2)
        ctk.CTkCheckBox(opts_inner, text="Abrir pasta do Drive após conclusão", variable=self.open_after).pack(anchor=tk.W, pady=2)

        # Botões
        button_frame = ctk.CTkFrame(main, fg_color="transparent")
        button_frame.pack(fill=tk.X, pady=(10, 0))
        ctk.CTkButton(button_frame, text="❌ Cancelar", command=self.window.destroy,
                      fg_color="gray40", hover_color="gray30", width=120).pack(side=tk.LEFT)
        ctk.CTkButton(button_frame, text="📂 Abrir Pasta Local", command=self.open_local_folder,
                      fg_color="gray40", hover_color="gray30", width=150).pack(side=tk.LEFT, padx=(10, 0))
        ctk.CTkButton(button_frame, text="📤 Enviar para Drive", command=self.start_upload,
                      fg_color=ACCENT, hover_color="#0088AA", width=160,
                      font=("Segoe UI", 11, "bold")).pack(side=tk.RIGHT)
    
    def select_drive_folder(self):
        """Seleciona pasta do Google Drive"""
        initial = self.drive_path.get() or self.app.last_dir
        folder = filedialog.askdirectory(
            title="Selecionar Pasta do Google Drive",
            initialdir=initial
        )
        if folder:
            self.drive_path.set(normalize_path(folder))
    
    def open_local_folder(self):
        """Abre pasta local no explorador"""
        try:
            if platform.system() == 'Windows':
                os.startfile(self.source_folder)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', self.source_folder])
            else:  # Linux
                subprocess.Popen(['xdg-open', self.source_folder])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {e}")
    
    def start_upload(self):
        """Inicia o processo de upload"""
        drive_path = self.drive_path.get().strip()
        
        if not drive_path:
            messagebox.showwarning("Aviso", "Selecione a pasta de destino no Google Drive!")
            return
        
        if not os.path.exists(drive_path) or not os.path.isdir(drive_path):
            messagebox.showerror("Erro", "Pasta de destino não encontrada!")
            return
        
        # Confirmar
        confirm_msg = f"Confirmar envio de {self.file_summary['total_files']} arquivo(s) para:\n\n{drive_path}\n\n"
        if not self.keep_local.get():
            confirm_msg += "⚠️ ATENÇÃO: Arquivos locais serão REMOVIDOS após o upload!\n\n"
        confirm_msg += "Deseja continuar?"
        
        if not messagebox.askyesno("Confirmar Upload", confirm_msg):
            return
        
        # Criar backup se solicitado
        if self.create_backup.get():
            try:
                self.create_backup_zip()
            except Exception as e:
                if not messagebox.askyesno("Erro no Backup", 
                    f"Erro ao criar backup: {e}\n\nContinuar mesmo assim?"):
                    return
        
        # Fechar janela atual
        self.window.destroy()
        
        # Abrir janela de progresso e iniciar upload
        options = {
            'keep_local': self.keep_local.get(),
            'open_after': self.open_after.get()
        }
        
        self.app.upload_to_drive(self.source_folder, drive_path, options)
    
    def create_backup_zip(self):
        """Cria backup em ZIP da pasta de saída"""
        import zipfile
        
        backup_name = f"backup_{os.path.basename(self.source_folder)}_{time.strftime('%Y%m%d_%H%M%S')}.zip"
        backup_path = os.path.join(os.path.dirname(self.source_folder), backup_name)
        
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(self.source_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, self.source_folder)
                    zipf.write(file_path, arcname)
        
        self.app.write_log(f"✓ Backup criado: {backup_name}")


class UploadProgressDialog:
    """Janela de progresso durante upload"""

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.cancelled = False
        self.paused = False

        self.window = ctk.CTkToplevel(parent)
        self.window.title("Enviando para Google Drive...")
        self.window.geometry("650x380")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - 325
        y = (self.window.winfo_screenheight() // 2) - 190
        self.window.geometry(f"+{x}+{y}")

        self.setup_ui()

    def setup_ui(self):
        ACCENT = "#00A8CC"
        main = ctk.CTkFrame(self.window, fg_color="transparent")
        main.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        ctk.CTkLabel(main, text="📤 Enviando arquivos para Google Drive",
                     font=("Segoe UI", 14, "bold"), text_color=ACCENT).pack(pady=(0, 20))

        self.status_label = ctk.CTkLabel(main, text="Preparando upload...", font=("Segoe UI", 11))
        self.status_label.pack(pady=(0, 15))

        self.progress = ctk.CTkProgressBar(main, width=550, height=18, mode='determinate')
        self.progress.pack(pady=(0, 8))
        self.progress.set(0)

        self.percent_label = ctk.CTkLabel(main, text="0%",
                                          font=("Segoe UI", 10, "bold"), text_color=ACCENT)
        self.percent_label.pack()

        self.current_file = ctk.CTkLabel(main, text="", font=("Segoe UI", 9), text_color="gray")
        self.current_file.pack(pady=(15, 5))

        self.stats_label = ctk.CTkLabel(main, text="0 / 0 arquivos • 0 MB / 0 MB", font=("Segoe UI", 9))
        self.stats_label.pack(pady=(0, 5))

        self.time_label = ctk.CTkLabel(main, text="Calculando tempo restante...",
                                       font=("Segoe UI", 9), text_color="gray")
        self.time_label.pack()

        button_frame = ctk.CTkFrame(main, fg_color="transparent")
        button_frame.pack(pady=(25, 0))

        self.cancel_btn = ctk.CTkButton(button_frame, text="❌ Cancelar", command=self.cancel,
                                        fg_color="gray40", hover_color="gray30", width=120)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def update_progress(self, current, total, current_file, bytes_sent, bytes_total, elapsed_time):
        """Atualiza o progresso do upload"""
        if self.cancelled:
            return False

        try:
            percent = (current / total) if total > 0 else 0
            self.progress.set(percent)
            self.percent_label.configure(text=f"{percent*100:.1f}%")
            self.status_label.configure(text=f"Enviando arquivo {current} de {total}...")
            self.current_file.configure(text=f"📄 {os.path.basename(current_file)}")

            mb_sent = bytes_sent / (1024 * 1024)
            mb_total = bytes_total / (1024 * 1024)
            self.stats_label.configure(
                text=f"{current} / {total} arquivos • {mb_sent:.1f} MB / {mb_total:.1f} MB")

            if current > 0 and elapsed_time > 0:
                remaining_time = (elapsed_time / current) * (total - current)
                if remaining_time < 60:
                    time_str = f"~{int(remaining_time)}s restantes"
                elif remaining_time < 3600:
                    time_str = f"~{int(remaining_time / 60)}m restantes"
                else:
                    time_str = f"~{int(remaining_time / 3600)}h restantes"
                self.time_label.configure(text=time_str)

            self.window.update()
            return True

        except Exception as e:
            print(f"Erro ao atualizar progresso: {e}")
            return True

    def cancel(self):
        """Cancela o upload"""
        if messagebox.askyesno("Cancelar Upload",
                               "Tem certeza que deseja cancelar o upload?\n\nArquivos já enviados permanecerão no Drive."):
            self.cancelled = True
            self.status_label.configure(text="❌ Cancelando...")
            self.cancel_btn.configure(state='disabled')
    
    def on_closing(self):
        """Intercepta fechamento da janela"""
        self.cancel()
    
    def close(self):
        """Fecha a janela"""
        try:
            self.window.destroy()
        except:
            pass


class UploadCompleteDialog:
    """Relatório final após upload"""

    def __init__(self, parent, app, results):
        self.parent = parent
        self.app = app
        self.results = results

        title = "Upload Concluído" if results['success'] > 0 else "Upload com Problemas"
        self.window = ctk.CTkToplevel(parent)
        self.window.title(title)
        self.window.geometry("700x600")
        self.window.transient(parent)
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - 350
        y = (self.window.winfo_screenheight() // 2) - 300
        self.window.geometry(f"+{x}+{y}")

        self.setup_ui()

    def setup_ui(self):
        ACCENT = "#00A8CC"
        main = ctk.CTkFrame(self.window, fg_color="transparent")
        main.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        if self.results['errors'] == 0:
            icon_text, title_text, color = "✅", "Upload Concluído com Sucesso!", "#4CAF50"
        else:
            icon_text, title_text, color = "⚠️", "Upload Concluído com Avisos", "#FF9800"

        ctk.CTkLabel(main, text=icon_text, font=("Segoe UI", 48)).pack()
        ctk.CTkLabel(main, text=title_text, font=("Segoe UI", 16, "bold"), text_color=color).pack(pady=(0, 20))

        stats_frame = ctk.CTkFrame(main, corner_radius=8)
        stats_frame.pack(fill=tk.X, pady=(0, 15))
        ctk.CTkLabel(stats_frame, text="📊 Estatísticas",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 5))
        stats_text = (f"✓ {self.results['success']} arquivo(s) enviado(s) com sucesso\n"
                      f"✗ {self.results['errors']} erro(s)\n"
                      f"⏱️ Tempo total: {self.results['duration']}\n"
                      f"💾 Dados transferidos: {self.results['size_mb']} MB\n"
                      f"🔗 Destino: {os.path.basename(self.results['drive_url'])}")
        ctk.CTkLabel(stats_frame, text=stats_text, justify=tk.LEFT).pack(anchor=tk.W, padx=15, pady=(0, 10))

        if self.results['errors'] > 0 and self.results.get('error_list'):
            error_frame = ctk.CTkFrame(main, corner_radius=8)
            error_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
            ctk.CTkLabel(error_frame, text="⚠️ Arquivos com Erro",
                         font=("Segoe UI", 10, "bold"), text_color="#FF9800").pack(anchor=tk.W, padx=15, pady=(10, 5))
            error_text = ctk.CTkTextbox(error_frame, height=180, font=("Consolas", 9))
            error_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
            for idx, error in enumerate(self.results['error_list'], 1):
                error_text.insert("end", f"{idx}. {os.path.basename(error['file'])}\n")
                error_text.insert("end", f"   Erro: {error['error']}\n\n")
            error_text.configure(state='disabled')

        action_frame = ctk.CTkFrame(main, fg_color="transparent")
        action_frame.pack(pady=(10, 0))
        ctk.CTkButton(action_frame, text="🔗 Abrir no Drive",
                      command=lambda: self.open_drive(self.results['drive_url']),
                      fg_color="gray40", hover_color="gray30", width=140).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(action_frame, text="📄 Salvar Relatório", command=self.save_report,
                      fg_color="gray40", hover_color="gray30", width=140).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(action_frame, text="✓ Fechar", command=self.window.destroy,
                      fg_color=ACCENT, hover_color="#0088AA", width=120,
                      font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=5)
    
    def open_drive(self, drive_url):
        """Abre pasta no explorador"""
        try:
            if platform.system() == 'Windows':
                os.startfile(drive_url)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', drive_url])
            else:
                subprocess.Popen(['xdg-open', drive_url])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {e}")
    
    def save_report(self):
        """Salva relatório em arquivo"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Arquivo de Texto", "*.txt")],
                initialfile=f"relatorio_upload_{time.strftime('%Y%m%d_%H%M%S')}.txt"
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("="*80 + "\n")
                    f.write("RELATÓRIO DE UPLOAD PARA GOOGLE DRIVE\n")
                    f.write("="*80 + "\n")
                    f.write(f"Data/Hora: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write(f"Arquivos enviados: {self.results['success']}\n")
                    f.write(f"Erros: {self.results['errors']}\n")
                    f.write(f"Tempo total: {self.results['duration']}\n")
                    f.write(f"Tamanho: {self.results['size_mb']} MB\n")
                    f.write(f"Destino: {self.results['drive_url']}\n")
                    f.write("="*80 + "\n\n")
                    
                    if self.results.get('error_list'):
                        f.write("ERROS:\n")
                        f.write("-"*80 + "\n")
                        for idx, error in enumerate(self.results['error_list'], 1):
                            f.write(f"{idx}. {error['file']}\n")
                            f.write(f"   Erro: {error['error']}\n\n")
                
                messagebox.showinfo("Sucesso", "Relatório salvo com sucesso!")
        
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar relatório: {e}")


# ==================== FUNÇÕES AUXILIARES ====================

def normalize_account(conta):
    """Normaliza conta removendo caracteres. Ex: '52938-2' -> '529382'"""
    if conta is None:
        return ""
    return re.sub(r'[^0-9]', '', str(conta))


def extract_credited_account_section(text):
    if not text:
        return ""
    
    # Padrões possíveis de cabeçalho da seção (variações)
    section_patterns = [
        r'dados\s+da\s+conta\s+creditada',
        r'conta\s+creditada',
        r'favorecido',
        r'benefici[aá]rio',
    ]
    
    # Padrões que indicam o fim da seção (início da próxima seção)
    end_patterns = [
        r'dados\s+do\s+pagador',
        r'dados\s+da\s+transfer[eê]ncia',
        r'dados\s+do\s+comprovante',
        r'autenticac[aã]o',
        r'valor',
        r'data\s+da\s+operac[aã]o',
    ]
    
    # Normalizar texto para busca (manter pontuação para melhor detecção)
    text_upper = text.upper()
    
    # Procurar início da seção
    start_pos = -1
    matched_pattern = None
    
    for pattern in section_patterns:
        match = re.search(pattern, text_upper, re.IGNORECASE)
        if match:
            start_pos = match.start()
            matched_pattern = pattern
            break
    
    # Se não encontrou a seção, retornar texto vazio
    if start_pos == -1:
        return ""
    
    # Procurar fim da seção (próxima seção ou fim razoável)
    end_pos = len(text)
    
    # Buscar a partir do início da seção encontrada
    text_after_start = text_upper[start_pos:]
    
    for pattern in end_patterns:
        # Buscar após o cabeçalho (pular pelo menos 20 caracteres para não pegar o próprio cabeçalho)
        match = re.search(pattern, text_after_start[50:], re.IGNORECASE)
        if match:
            # Ajustar posição relativa ao texto original
            candidate_end = start_pos + 50 + match.start()
            if candidate_end < end_pos:
                end_pos = candidate_end
            break
    
    # Se não encontrou fim explícito, limitar a um tamanho razoável (ex: 500 caracteres)
    if end_pos == len(text):
        end_pos = min(start_pos + 500, len(text))
    
    # Extrair seção
    section_text = text[start_pos:end_pos]
    
    return section_text


def extract_auth_code(text):
    """
    Extrai o código de autenticação do comprovante.
    Exemplos Itaú: sequência hexadecimal de 32-40 chars após 'Autenticação:'.
    Retorna a string do código em maiúsculas, ou None se não encontrar.
    """
    if not text:
        return None
    # Padrão: após 'Autenticação' (com ou sem acento, com ou sem ':'), captura sequência hex longa
    match = re.search(
        r'Autentica[çc][aã]o\s*:?\s*([A-Fa-f0-9]{20,})',
        text,
        re.IGNORECASE
    )
    if match:
        return match.group(1).upper()
    # Fallback: qualquer sequência hex isolada com 32+ chars (CTRL codes, etc.)
    match = re.search(r'(?<![A-Fa-f0-9])([A-Fa-f0-9]{32,})(?![A-Fa-f0-9])', text)
    if match:
        return match.group(1).upper()
    return None


def extract_pdf_pages(pdf_path):
    """Extrai texto de cada página do PDF"""
    pages = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            # Texto normalizado para busca: remove acentos, converte para maiúsculas e colapsa espaços
            def normalize_search_text(s):
                if not s:
                    return ""
                nf = unicodedata.normalize('NFKD', s)
                ascii_s = nf.encode('ascii', 'ignore').decode('ascii')
                # manter apenas letras, dígitos e espaços
                cleaned = re.sub(r'[^A-Za-z0-9\s]', ' ', ascii_s)
                cleaned = re.sub(r'\s+', ' ', cleaned).strip().upper()
                return cleaned

            # Extrair seção específica "Dados da Conta Creditada"
            credited_section = extract_credited_account_section(text)
            # Código de autenticação único do comprovante
            auth_code = extract_auth_code(text)
            
            pages[i] = {
                'text': text,
                'numbers': normalize_account(text),
                'norm_text': normalize_search_text(text),
                # Novos campos para busca na seção específica
                'credited_section': credited_section,
                'credited_numbers': normalize_account(credited_section),
                'credited_norm_text': normalize_search_text(credited_section),
                # Código de autenticação para deduplicação
                'auth_code': auth_code
            }
    return pages


def extract_name_from_page(page_data):
    """
    Extrai o nome do beneficiário/creditado diretamente do texto da página do PDF.
    Busca o campo 'Nome:' dentro da seção 'Dados da Conta Creditada'.
    Retorna o nome encontrado (str) ou None se não encontrar.
    """
    text = page_data.get('credited_section', '') or page_data.get('text', '')
    if not text:
        return None

    # Padrões para capturar o nome após 'Nome:' (e variantes)
    patterns = [
        r'Nome\s*:\s*([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú\s\.\-]{2,80}?)(?:\n|Agência|Ag[\.\:]|Conta|CPF|CNPJ|$)',
        r'Nome\s*:\s*(.+?)(?:\n|Agência|Ag[\.\:]|Conta|CPF|CNPJ)',
        r'Nome\s*:\s*(.+)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            nome = match.group(1).strip()
            # Filtrar resultados muito curtos ou que sejam apenas dígitos
            if nome and len(nome) >= 3 and not nome.isdigit():
                # Remover traços, pontos e espaços duplicados finais
                nome = re.sub(r'\s+', ' ', nome).strip()
                return nome

    return None


def find_account_pages(conta, agencia, pages):
    """
    Busca páginas onde TANTO a conta QUANTO a agência aparecem juntos NA SEÇÃO 'DADOS DA CONTA CREDITADA'.
    Se não encontrar, tenta com os valores invertidos (conta<->agência) caso estejam trocados na planilha.
    Como último recurso, faz busca ampla procurando qualquer um dos valores.
    Retorna tupla: (lista_de_páginas, invertido) onde invertido=True se usou valores trocados.
    """
    found = []
    conta_norm = normalize_account(conta)
    agencia_norm = normalize_account(agencia)
    
    if not conta_norm or len(conta_norm) < 3:
        return found, False
    
    if not agencia_norm or len(agencia_norm) < 3:
        return found, False
    
    # Função auxiliar para buscar número exato com delimitadores
    def find_exact_number(number, text):
        """
        Busca número exato no texto, garantindo que não é parte de outro número.
        O número deve ser exatamente igual ao que está na planilha.
        """
        if not number or not text:
            return False
        
        # Criar padrão que permite separadores entre dígitos mas exige delimitadores nas bordas
        digits = list(number)
        # Padrão: início ou não-dígito, depois os dígitos (com possíveis separadores), depois fim ou não-dígito  
        # (?:[\s\-\.]*\d)? permite um dígito verificador opcional no final
        pattern = r'(?<!\d)' + r'[\s\-\.]*'.join(digits) + r'(?:[\s\-\.]*\d)?(?!\d)'
        try:
            if re.search(pattern, text):
                return True
        except re.error:
            pass
        return False
    
    def buscar_com_valores(val_conta, val_agencia):
        """Busca páginas com os valores de conta e agência fornecidos"""
        resultados = []
        
        for num, data in pages.items():
            # Usar dados da seção "Dados da Conta Creditada" (se existir)
            credited_section = data.get('credited_section', '')
            
            # Se não encontrou a seção, pular esta página
            if not credited_section or len(credited_section) < 20:
                continue
            
            tem_conta = False
            tem_agencia = False
            
            # Verifica se tem a conta NA SEÇÃO CREDITADA (busca exata)
            if val_conta and find_exact_number(val_conta, credited_section):
                tem_conta = True
            
            # Busca alternativa: sem dígito verificador (último recurso)
            if not tem_conta and len(val_conta) > 4:
                conta_sem_dv = val_conta[:-1]
                if len(conta_sem_dv) >= 4 and find_exact_number(conta_sem_dv, credited_section):
                    tem_conta = True
            
            # Verifica se tem a agência NA SEÇÃO CREDITADA (busca exata)
            if val_agencia and find_exact_number(val_agencia, credited_section):
                tem_agencia = True
            
            # SÓ adiciona se encontrou AMBOS: conta E agência
            # Adiciona se encontrou ao menos a conta (agência é opcional)
            if tem_conta:
                if num not in resultados:
                    resultados.append(num)
        
        return resultados
    
    # Primeira tentativa: valores originais (conta na coluna conta, agência na coluna agência)
    found = buscar_com_valores(conta_norm, agencia_norm)
    
    if found:
        return found, False  # Encontrou com valores originais
    
    # Segunda tentativa: valores INVERTIDOS (conta<->agência trocados na planilha)
    # Só tenta se os valores forem diferentes entre si
    if conta_norm != agencia_norm:
        found_invertido = buscar_com_valores(agencia_norm, conta_norm)
        if found_invertido:
            return found_invertido, True  # Encontrou com valores invertidos
    
    # Terceira tentativa: BUSCA TEXTUAL AMPLA (qualquer um dos valores em qualquer lugar)
    # Para casos onde os dados estão em colunas erradas ou em branco
    found_ampla = []
    for num, data in pages.items():
        credited_section = data.get('credited_section', '')
        
        if not credited_section or len(credited_section) < 20:
            continue
        
        # Buscar QUALQUER UM dos valores (conta OU agência) em QUALQUER LUGAR da seção
        encontrou_algum = False
        
        # Tentar encontrar conta
        if conta_norm and find_exact_number(conta_norm, credited_section):
            encontrou_algum = True
        
        # Tentar encontrar agência
        if not encontrou_algum and agencia_norm and find_exact_number(agencia_norm, credited_section):
            encontrou_algum = True
        
        # Busca sem dígito verificador (último recurso)
        if not encontrou_algum:
            if len(conta_norm) > 4:
                conta_sem_dv = conta_norm[:-1]
                if len(conta_sem_dv) >= 4 and find_exact_number(conta_sem_dv, credited_section):
                    encontrou_algum = True
            
            if not encontrou_algum and len(agencia_norm) > 4:
                agencia_sem_dv = agencia_norm[:-1]
                if len(agencia_sem_dv) >= 4 and find_exact_number(agencia_sem_dv, credited_section):
                    encontrou_algum = True
        
        if encontrou_algum and num not in found_ampla:
            found_ampla.append(num)
    
    if found_ampla:
        return found_ampla, False  # Encontrou com busca ampla
    
    return found, False


def create_pdf(pdf_path, page_numbers, output_path):
    """Cria PDF com páginas específicas"""
    if not page_numbers:
        return 0

    reader = None
    writer = None

    try:
        # Abrir o arquivo PDF fonte
        reader = PyPDF2.PdfReader(pdf_path)

        # Criar um novo writer para cada arquivo
        writer = PyPDF2.PdfWriter()

        # Adicionar apenas as páginas especificadas
        pages_added = 0
        for num in page_numbers:
            if 0 <= num < len(reader.pages):
                page = reader.pages[num]
                writer.add_page(page)
                pages_added += 1

        # Verificar se há páginas e salvar
        if pages_added > 0:
            # Garantir que NÃO sobrescrevemos arquivos já existentes
            target = output_path
            if os.path.exists(target):
                base, ext = os.path.splitext(target)
                # tentar com sufixo timestamp
                stamp = str(int(time.time() * 1000))
                candidate = f"{base}_{stamp}{ext}"
                # em casos raros de colisão, iterar
                i = 1
                while os.path.exists(candidate):
                    candidate = f"{base}_{stamp}_{i}{ext}"
                    i += 1
                target = candidate

            # Salvar diretamente no arquivo de destino
            try:
                with open(target, 'wb') as out:
                    writer.write(out)
            except Exception as e:
                print(f"Erro ao salvar PDF {target}: {e}")
                return 0

            # Retornar número de páginas efetivamente escritas
            return pages_added

        # Nenhuma página válida para escrever
        return 0

    except Exception as e:
        print(f"Erro criar PDF: {e}")
        return 0

    finally:
        # Limpar referências
        writer = None
        reader = None


def normalize_path(path):
    """Normaliza path garantindo encoding correto para Windows/OneDrive/Google Drive"""
    if not path:
        return path
    
    try:
        # Converter para string se necessário
        if isinstance(path, bytes):
            path = path.decode('utf-8', errors='replace')
        
        path = str(path).strip()
        
        # Normalizar barras para o sistema operacional
        if platform.system() == 'Windows':
            path = path.replace('/', '\\')
        
        # Resolver Path para garantir formato correto
        path_obj = Path(path)
        # Usar resolve() para expandir caminhos relativos e normalizar
        try:
            resolved = path_obj.resolve()
            return str(resolved)
        except (OSError, RuntimeError):
            # Se resolve() falhar, retornar path normalizado básico
            return os.path.normpath(path)
    except Exception:
        # Fallback: retornar path original
        return path


def clean_filename(name):
    """Remove caracteres inválidos"""
    if not name or str(name).lower() == 'nan':
        return "sem_nome"
    name = str(name)
    for c in '<>:"/\\|?*\n\r\t':
        name = name.replace(c, '_')
    return ' '.join(name.split())[:100].strip()


def find_column(df, names):
    """Encontra coluna pelo nome - busca exata primeiro, depois parcial"""
    # Primeira passada: busca exata
    for col in df.columns:
        for name in names:
            if str(col).lower().strip() == name.lower().strip():
                return col
    
    # Segunda passada: busca parcial
    for col in df.columns:
        for name in names:
            if name.lower() in str(col).lower():
                return col
    return None


class App:
    ACCENT = "#00A8CC"
    ACCENT_HOVER = "#0088AA"

    def __init__(self, root):
        self.root = root
        self.root.title("PD7Lab - Extrator de Comprovantes PDF v1.0.0")
        self.root.geometry("950x780")
        self.root.minsize(850, 680)

        try:
            icon_path = resource_path("pd7-escudo.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self.pdf_folder_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.out_var = tk.StringVar(value="comprovantes_extraidos")
        self.df = None
        self.conta_col = None
        self.agencia_col = None
        self.nome_col = None
        self.ccusto_col = None
        self.last_dir = os.path.expanduser("~")

        self.force_reprocess_var = tk.BooleanVar(value=False)
        self.debug_mode_var = tk.BooleanVar(value=False)

        self.start_time = None
        self.timer_running = False
        self.timer_label = None

        self.logo_image = None
        self.logo_label = None

        self.current_theme = 'light'

        self.processed_pdfs_file = "pdfs_processados.json"
        self.processed_pdfs = self.load_processed_pdfs()

        self.last_output_folder = None
        self.last_process_stats = None

        self.setup_ui()
    
    def load_processed_pdfs(self):
        """Carrega lista de PDFs já processados"""
        try:
            if os.path.exists(self.processed_pdfs_file):
                with open(self.processed_pdfs_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}
    
    def save_processed_pdfs(self):
        """Salva lista de PDFs processados"""
        try:
            with open(self.processed_pdfs_file, 'w', encoding='utf-8') as f:
                json.dump(self.processed_pdfs, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar histórico: {e}")
    
    def get_pdf_fingerprint(self, pdf_path):
        """Gera identificador único para PDF (nome + tamanho + data modificação)"""
        try:
            stat = os.stat(pdf_path)
            return f"{os.path.basename(pdf_path)}_{stat.st_size}_{stat.st_mtime}"
        except:
            return None
    
    def toggle_theme(self):
        """Alterna entre tema claro e escuro"""
        self.current_theme = 'dark' if self.current_theme == 'light' else 'light'
        ctk.set_appearance_mode('dark' if self.current_theme == 'dark' else 'light')
        icon = "☀️ Modo Claro" if self.current_theme == 'dark' else "🌙 Modo Escuro"
        try:
            self.theme_btn.configure(text=icon)
        except Exception:
            pass
        theme_name = 'Escuro' if self.current_theme == 'dark' else 'Claro'
        self.write_log(f"🎨 Tema alterado para: {theme_name}")
    
    def setup_ui(self):
        ACCENT = self.ACCENT
        HOVER = self.ACCENT_HOVER

        # Container principal
        main = ctk.CTkFrame(self.root, fg_color="transparent")
        main.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Header
        header_frame = ctk.CTkFrame(main, fg_color="transparent")
        header_frame.pack(fill=tk.X, pady=(0, 12))

        try:
            if Image and ImageTk:
                logo_file = 'pd7lab-dark.jpeg' if self.current_theme == 'light' else 'pd7.png'
                logo_path = resource_path(logo_file)
                if os.path.exists(logo_path):
                    logo_img = Image.open(logo_path)
                    new_height = 60
                    new_width = int(new_height * logo_img.width / logo_img.height)
                    logo_img = logo_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    self.logo_image = ImageTk.PhotoImage(logo_img)
                    self.logo_label = ctk.CTkLabel(header_frame, image=self.logo_image, text="")
                    self.logo_label.pack(side=tk.LEFT, padx=(0, 15))
        except Exception as e:
            print(f"Logo loading warning: {e}")

        header_text = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ctk.CTkLabel(header_text, text="Extrator de Comprovantes PDF v1.0.0",
                     font=("Segoe UI", 18, "bold"), text_color=ACCENT).pack(anchor=tk.W)
        ctk.CTkLabel(header_text, text="Automatize a extração de comprovantes bancários",
                     font=("Segoe UI", 9), text_color="gray").pack(anchor=tk.W)

        theme_icon = "🌙 Modo Escuro" if self.current_theme == 'light' else "☀️ Modo Claro"
        self.theme_btn = ctk.CTkButton(header_frame, text=theme_icon, command=self.toggle_theme,
                                        width=140, fg_color="gray40", hover_color="gray30")
        self.theme_btn.pack(side=tk.RIGHT)

        # Separador
        ctk.CTkFrame(main, height=2, fg_color=ACCENT).pack(fill=tk.X, pady=(0, 12))

        # Arquivos
        files_frame = ctk.CTkFrame(main, corner_radius=10)
        files_frame.pack(fill=tk.X, pady=(0, 8))
        ctk.CTkLabel(files_frame, text="📁 Arquivos",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 4))

        grid = ctk.CTkFrame(files_frame, fg_color="transparent")
        grid.pack(fill=tk.X, padx=15, pady=(0, 10))
        grid.columnconfigure(1, weight=1)

        for row, (label, var, cmd, val_cmd) in enumerate([
            ("Pasta PDFs:",    self.pdf_folder_var, self.get_pdf_folder, self.validate_pdf_folder),
            ("Planilha Excel:", self.excel_var,       self.get_excel,      self.validate_excel),
            ("Pasta de Saída:", self.out_var,          self.get_out,        self.validate_out),
        ]):
            ctk.CTkLabel(grid, text=label).grid(row=row, column=0, sticky=tk.W, padx=(0, 12), pady=5)
            entry = ctk.CTkEntry(grid, textvariable=var, font=("Segoe UI", 10))
            entry.grid(row=row, column=1, sticky='ew', padx=(0, 10), pady=5)
            entry.bind('<Return>', lambda e, v=val_cmd: v())
            ctk.CTkButton(grid, text="Procurar...", width=110, command=cmd,
                          fg_color=ACCENT, hover_color=HOVER).grid(row=row, column=2, pady=5)

        # Timer
        timer_row = ctk.CTkFrame(main, corner_radius=8)
        timer_row.pack(fill=tk.X, pady=(8, 6))
        self.timer_label = ctk.CTkLabel(timer_row, text="⏱️ Tempo: 00:00:00.000",
                                         font=("Segoe UI", 10, "bold"), text_color=ACCENT)
        self.timer_label.pack(side=tk.LEFT, padx=15, pady=8)

        # Opções
        opts_frame = ctk.CTkFrame(main, corner_radius=10)
        opts_frame.pack(fill=tk.X, pady=(6, 8))
        ctk.CTkLabel(opts_frame, text="⚙️ Opções de Processamento",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 4))
        opts_inner = ctk.CTkFrame(opts_frame, fg_color="transparent")
        opts_inner.pack(fill=tk.X, padx=15, pady=(0, 10))
        try:
            ctk.CTkCheckBox(opts_inner, text="Ignorar histórico (forçar reprocessamento)",
                            variable=self.force_reprocess_var).pack(side=tk.LEFT, padx=(0, 15))
            ctk.CTkCheckBox(opts_inner, text="🔧 Debug",
                            variable=self.debug_mode_var).pack(side=tk.LEFT, padx=(0, 15))
            ctk.CTkButton(opts_inner, text="🗑️ Limpar Histórico",
                          command=self.clear_processed_history, width=150,
                          fg_color="gray40", hover_color="gray30").pack(side=tk.LEFT, padx=(0, 8))
            ctk.CTkButton(opts_inner, text="🔍 Buscar Não Encontrados",
                          command=self.search_missing, width=200,
                          fg_color=ACCENT, hover_color=HOVER).pack(side=tk.LEFT)
        except Exception:
            pass

        # Controles (botão processar + barra de progresso)
        controls = ctk.CTkFrame(main, fg_color="transparent")
        controls.pack(fill=tk.X, pady=(10, 6))
        self.controls_frame = controls

        self.btn = ctk.CTkButton(controls, text="▶  PROCESSAR COMPROVANTES",
                                  command=self.start, fg_color=ACCENT, hover_color=HOVER,
                                  font=("Segoe UI", 11, "bold"), height=42, corner_radius=8)
        self.btn.pack(side=tk.LEFT, padx=(0, 15))

        self.prog = ctk.CTkProgressBar(controls, mode='indeterminate', height=20)
        self.prog.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))
        self.prog.set(0)

        self.status_var = tk.StringVar(value="Pronto")
        self.status_label = ctk.CTkLabel(controls, text="Pronto",
                                          font=("Segoe UI", 9), text_color="gray")
        self.status_label.pack(side=tk.LEFT)
        self.status_var.trace_add('write', lambda *_: self.status_label.configure(text=self.status_var.get()))

        # Log
        log_frame = ctk.CTkFrame(main, corner_radius=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        ctk.CTkLabel(log_frame, text="📋 Log de Processamento",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=15, pady=(10, 4))
        self.log = ctk.CTkTextbox(log_frame, font=("Consolas", 9), state='disabled',
                                   wrap='word', border_width=1)
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
    
    def update_timer(self):
        """Atualiza o cronômetro a cada 100ms"""
        if self.timer_running and self.start_time:
            elapsed = time.time() - self.start_time
            hours, remainder = divmod(int(elapsed), 3600)
            minutes, seconds = divmod(remainder, 60)
            milliseconds = int((elapsed % 1) * 1000)
            time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"
            self.timer_label.configure(text=f"⏱️ Tempo: {time_str}")
            self.root.after(100, self.update_timer)
    
    def start_timer(self):
        """Inicia o cronômetro"""
        self.start_time = time.time()
        self.timer_running = True
        self.timer_label.configure(text="⏱️ Tempo: 00:00:00.000")
        self.update_timer()
    
    def stop_timer(self):
        """Para o cronômetro e retorna tempo decorrido"""
        self.timer_running = False
        if self.start_time:
            elapsed = time.time() - self.start_time
            return elapsed
        return 0
    
    def format_time(self, seconds):
        """Formata segundos para formato legível com milissegundos"""
        hours, remainder = divmod(int(seconds), 3600)
        minutes, secs = divmod(remainder, 60)
        milliseconds = int((seconds % 1) * 1000)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}.{milliseconds:03d}"
    
    def get_pdf_folder(self):
        """Seleciona pasta usando explorador nativo do SO"""
        try:
            folder = self._native_select_folder("Selecionar Pasta com PDFs de Comprovantes")
            if folder:
                # Normalizar path para corrigir problemas de encoding
                folder = normalize_path(folder)
                
                # Verificar se a pasta existe após normalização
                if not os.path.exists(folder):
                    self.write_log(f"⚠️ Pasta não encontrada após normalização: {folder}")
                    messagebox.showerror("Erro", f"Pasta não encontrada: {folder}")
                    return
                
                if not os.path.isdir(folder):
                    self.write_log(f"⚠️ Caminho não é uma pasta: {folder}")
                    messagebox.showerror("Erro", f"Caminho não é uma pasta válida")
                    return
                
                self.pdf_folder_var.set(folder)
                self.last_dir = folder
                
                # Usar múltiplos métodos para contar PDFs (compatível com OneDrive)
                pdf_count = 0
                try:
                    counts = {}
                    
                    # Método 1: os.listdir
                    try:
                        count1 = len([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                        counts['listdir'] = count1
                    except Exception as e1:
                        self.write_log(f"  ⚠️ listdir falhou: {e1}")
                        counts['listdir'] = 0
                    
                    # Método 2: Path.iterdir (mais confiável)
                    try:
                        path_obj = Path(folder)
                        count2 = len([f for f in path_obj.iterdir() if f.is_file() and f.suffix.lower() == '.pdf'])
                        counts['iterdir'] = count2
                    except Exception as e2:
                        self.write_log(f"  ⚠️ iterdir falhou: {e2}")
                        counts['iterdir'] = 0
                    
                    # Método 3: os.scandir (eficiente)
                    try:
                        with os.scandir(folder) as entries:
                            count3 = len([e for e in entries if e.is_file() and e.name.lower().endswith('.pdf')])
                        counts['scandir'] = count3
                    except Exception as e3:
                        self.write_log(f"  ⚠️ scandir falhou: {e3}")
                        counts['scandir'] = 0
                    
                    pdf_count = max(counts.values()) if counts else 0
                    self.write_log(f"✓ Pasta PDFs: {os.path.basename(folder)} ({pdf_count} PDFs)")
                    
                    # Mostrar diferenças nos métodos se houver
                    if len(set(counts.values())) > 1:
                        methods_str = ", ".join([f"{k}={v}" for k, v in counts.items()])
                        self.write_log(f"  ℹ️ Métodos: {methods_str}")
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao contar PDFs: {e}")
                    self.write_log(f"  Pasta: {folder}")
            else:
                return
        except Exception as e:
            self.write_log(f"❌ Erro ao selecionar pasta: {e}")
            messagebox.showerror("Erro", f"Erro ao selecionar pasta: {e}")
    
    def get_excel(self):
        """Seleciona arquivo Excel usando explorador nativo do SO"""
        try:
            arquivo = self._native_select_file("Selecionar Planilha Excel", [("Todos os arquivos", "*.*")])
            if arquivo:
                # Normalizar path
                arquivo = normalize_path(arquivo)
                
                if os.path.isfile(arquivo):
                    self.excel_var.set(arquivo)
                    self.last_dir = os.path.dirname(arquivo)
                    self.write_log(f"✓ Excel: {os.path.basename(arquivo)}")
                    self.load_excel(arquivo)
                else:
                    self.write_log("⚠️ Arquivo selecionado não existe.")
                    messagebox.showwarning("Arquivo inválido", "O arquivo selecionado não existe.")
            else:
                return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar Excel: {e}")
    
    def load_excel(self, path):
        try:
            xl = pd.ExcelFile(path)
            sheet_names = xl.sheet_names
            all_rows = []
            sheets_loaded = 0

            def _has_data(val):
                """Retorna True se o valor tem conteúdo real (não vazio/nan/traço)."""
                if val is None:
                    return False
                s = str(val).strip()
                return s not in ('', '-', 'nan', 'NaN', 'None', 'NaT', 'nat')

            for sheet in sheet_names:
                # Leitura inicial para detectar nomes das colunas
                df_sheet = pd.read_excel(path, sheet_name=sheet)
                if df_sheet.empty:
                    continue

                nome_col = find_column(df_sheet, ['nome social', 'nome', 'funcionario'])
                if not nome_col:
                    continue

                agencia_col  = find_column(df_sheet, ['agencia', 'agência', 'ag', 'agency'])
                conta_col    = find_column(df_sheet, ['conta', 'account', 'conta corrente'])
                ccusto_col   = find_column(df_sheet, [
                    'descrição ccusto', 'descricao ccusto', 'descrição de ccusto',
                    'descricao de ccusto', 'desc ccusto', 'ccusto', 'centro de custo', 'setor'
                ])

                # Detectar coluna duplicada Conta.1 (segunda coluna Conta no export do RH)
                # pandas renomeia automaticamente duplicatas: Conta → Conta, Conta → Conta.1
                conta1_col = 'Conta.1' if 'Conta.1' in df_sheet.columns else None

                # Reler forçando colunas bancárias como texto (preserva zeros à esquerda)
                dtype_dict = {c: str for c in [conta_col, agencia_col, conta1_col] if c}
                df_sheet = pd.read_excel(path, sheet_name=sheet, dtype=dtype_dict)

                sheet_ccusto = sheet.strip()
                rows_in_sheet = 0

                for _, row in df_sheet.iterrows():
                    nome_val = row.get(nome_col)
                    if pd.isna(nome_val) or str(nome_val).strip() == '':
                        continue

                    # ── Lógica de banco (dois formatos coexistem na planilha) ──────────
                    # Formato A: coluna Agencia preenchida  → Agencia=agência, Conta=conta
                    # Formato B: coluna Agencia vazia       → Conta=agência,  Conta.1=conta
                    agencia_raw = str(row.get(agencia_col, '')) if agencia_col else ''
                    conta_raw   = str(row.get(conta_col, ''))   if conta_col   else ''
                    conta1_raw  = str(row.get(conta1_col, ''))  if conta1_col  else ''

                    if _has_data(agencia_raw):
                        agencia_val = agencia_raw.strip()
                        conta_val   = conta_raw.strip()
                    elif _has_data(conta1_raw):
                        # Conta.1 presente → Conta guarda agência, Conta.1 guarda a conta
                        agencia_val = conta_raw.strip()
                        conta_val   = conta1_raw.strip()
                    else:
                        agencia_val = ''
                        conta_val   = conta_raw.strip()

                    # ── Centro de custo ───────────────────────────────────────────────
                    if ccusto_col and _has_data(row.get(ccusto_col)):
                        ccusto_val = str(row[ccusto_col]).strip()
                    else:
                        ccusto_val = sheet_ccusto  # nome da aba como fallback

                    all_rows.append({
                        '_nome':    str(nome_val).strip(),
                        '_agencia': agencia_val,
                        '_conta':   conta_val,
                        '_ccusto':  ccusto_val or sheet_ccusto,
                    })
                    rows_in_sheet += 1

                if rows_in_sheet > 0:
                    self.write_log(f"  ✓ Aba '{sheet}': {rows_in_sheet} registros")
                    sheets_loaded += 1

            if not all_rows:
                self.write_log("⚠️ Nenhum registro válido encontrado na planilha!")
                return

            self.df = pd.DataFrame(all_rows)
            self.conta_col   = '_conta'
            self.agencia_col = '_agencia'
            self.nome_col    = '_nome'
            self.ccusto_col  = '_ccusto'

            # Estatísticas de centros de custo detectados
            ccusto_counts = self.df['_ccusto'].value_counts()
            sem_banco = (self.df['_conta'].str.strip() == '').sum()

            self.write_log(f"\nAbas carregadas: {sheets_loaded} | Total de registros: {len(self.df)}")
            self.write_log(f"  📊 Centros de custo (CCusto): {len(ccusto_counts)} únicos")
            for ccusto_name, count in ccusto_counts.items():
                self.write_log(f"     • {ccusto_name}: {count} funcionários")
            if sem_banco:
                self.write_log(f"  ⚠️ {sem_banco} registros sem dados bancários (serão ignorados)")
            self.write_log(f"✓ Colunas mapeadas: Nome, Agência+Conta (duplo formato), Descrição Ccusto")
        except Exception as e:
            self.write_log(f"Erro ao carregar planilha: {e}")
            import traceback
            self.write_log(traceback.format_exc())
    
    def get_out(self):
        """Seleciona pasta de saída usando explorador nativo do SO"""
        try:
            folder = self._native_select_folder("Selecionar Pasta de Saída")
            if folder:
                # Normalizar path
                folder = normalize_path(folder)
                self.out_var.set(folder)
                self.last_dir = folder
                self.write_log(f"✓ Pasta de saída: {folder}")
            else:
                return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar pasta: {e}")
    
    def _native_select_folder(self, title):
        folder = filedialog.askdirectory(initialdir=self.last_dir, title=title)
        if folder:
            return normalize_path(folder)
        return None
    
    def _native_select_file(self, title, filetypes):
        arquivo = filedialog.askopenfilename(initialdir=self.last_dir, title=title, filetypes=filetypes)
        if arquivo:
            return normalize_path(arquivo)
        return None
    
    def validate_pdf_folder(self):
        path = normalize_path(self.pdf_folder_var.get().strip())
        if path and os.path.exists(path) and os.path.isdir(path):
            self.last_dir = path
            try:
                pdf_count_listdir = len([f for f in os.listdir(path) if f.lower().endswith('.pdf')])
                path_obj = Path(path)
                pdf_count_iterdir = len([f for f in path_obj.iterdir() if f.is_file() and f.suffix.lower() == '.pdf'])
                pdf_count = max(pdf_count_listdir, pdf_count_iterdir)
                self.write_log(f"✓ Pasta PDFs: {os.path.basename(path)} ({pdf_count} PDFs)")
                if pdf_count_listdir != pdf_count_iterdir:
                    self.write_log(f"  ℹ️ Métodos: listdir={pdf_count_listdir}, iterdir={pdf_count_iterdir}")
            except Exception as e:
                self.write_log(f"⚠️ Erro ao contar PDFs: {e}")
        elif path:
            messagebox.showwarning("Aviso", "Pasta não encontrada!")
    
    def validate_excel(self):
        path = normalize_path(self.excel_var.get().strip())
        if path and os.path.exists(path) and (path.endswith('.xlsx') or path.endswith('.xls')):
            self.last_dir = os.path.dirname(path)
            self.write_log(f"✓ Excel: {os.path.basename(path)}")
            self.load_excel(path)
        elif path:
            messagebox.showwarning("Aviso", "Arquivo Excel não encontrado!")
    
    def validate_out(self):
        path = self.out_var.get().strip()
        if path:
            self.write_log(f"✓ Pasta: {path}")
    
    def write_log(self, msg):
        try:
            self.log.configure(state='normal')
            self.log.insert("end", msg + "\n")
            self.log._textbox.see("end")
            self.log.configure(state='disabled')
            self.root.update()
        except Exception:
            print(msg)

    def clear_processed_history(self):
        """Apaga o histórico de PDFs processados (arquivo e memória)"""
        try:
            if messagebox.askyesno("Confirmar", "Tem certeza que deseja limpar o histórico de PDFs processados?"):
                self.processed_pdfs = {}
                try:
                    if os.path.exists(self.processed_pdfs_file):
                        os.remove(self.processed_pdfs_file)
                except Exception as e:
                    self.write_log(f"Erro ao limpar histórico: {e}")
                else:
                    self.write_log("✓ Histórico de PDFs processados limpo.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao limpar histórico: {e}")
    
    def search_missing(self):
        """Busca assistida para comprovantes não encontrados"""
        if not self.pdf_folder_var.get():
            messagebox.showwarning("Aviso", "Selecione a pasta de PDFs primeiro!")
            return
        
        # Perguntar origem dos dados
        choice_win = tk.Toplevel(self.root)
        choice_win.title("Origem dos Dados")
        choice_win.geometry("450x250")
        choice_win.resizable(False, False)
        
        # Centralizar janela
        choice_win.transient(self.root)
        choice_win.grab_set()
        
        frame = ttk.Frame(choice_win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="De onde deseja carregar os itens para buscar?", 
                 font=('Segoe UI', 11, 'bold')).pack(pady=(0, 20))
        
        result = {'source': None}
        
        def use_txt():
            result['source'] = 'txt'
            choice_win.destroy()
        
        def use_excel():
            result['source'] = 'excel'
            choice_win.destroy()
        
        def cancel():
            result['source'] = None
            choice_win.destroy()
        
        # Botão 1: Arquivo TXT
        btn_frame1 = ttk.Frame(frame)
        btn_frame1.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame1, text="📄 Arquivo TXT de Não Encontrados", 
                  command=use_txt, width=40).pack()
        ttk.Label(btn_frame1, text="Selecionar arquivo TXT gerado anteriormente", 
                 font=('Segoe UI', 8), foreground='gray').pack()
        
        # Separador
        ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=15)
        
        # Botão 2: Excel
        btn_frame2 = ttk.Frame(frame)
        btn_frame2.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame2, text="📊 Planilha Excel Completa", 
                  command=use_excel, width=40).pack()
        ttk.Label(btn_frame2, text="Buscar todos os registros do Excel", 
                 font=('Segoe UI', 8), foreground='gray').pack()
        
        # Botão cancelar
        ttk.Button(frame, text="Cancelar", command=cancel, width=15).pack(pady=(20, 0))
        
        # Aguardar escolha
        self.root.wait_window(choice_win)
        
        missing_items = []
        
        if result['source'] == 'txt':
            # Selecionar arquivo TXT
            txt_file = filedialog.askopenfilename(
                title="Selecionar arquivo de não encontrados",
                initialdir=self.last_dir,
                filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
            )
            
            if not txt_file:
                return
            
            txt_file = normalize_path(txt_file)
            missing_items = self.parse_missing_txt(txt_file)
            
            if not missing_items:
                messagebox.showinfo("Info", "Nenhum item encontrado no arquivo TXT.")
                return
            
            self.write_log(f"\n{'='*50}")
            self.write_log(f"🔍 BUSCA ASSISTIDA - Arquivo TXT")
            self.write_log(f"{'='*50}")
            self.write_log(f"📄 Arquivo: {os.path.basename(txt_file)}")
            self.write_log(f"📊 Total de itens: {len(missing_items)}")
            
        elif result['source'] == 'excel':
            # Usar Excel carregado ou solicitar
            if self.df is None or not self.conta_col or not self.nome_col or not self.ccusto_col:
                if not self.excel_var.get():
                    messagebox.showwarning("Aviso", "Carregue uma planilha Excel primeiro!")
                    return
                else:
                    messagebox.showwarning("Aviso", "Excel não está carregado corretamente.\nVerifique as colunas necessárias.")
                    return
            
            # Carregar todos os registros do Excel
            missing_items = []
            for row_idx, row in self.df.iterrows():
                conta = row[self.conta_col]
                nome = row[self.nome_col]
                ccusto = row[self.ccusto_col]
                
                if pd.isna(conta) or str(conta).strip() == '':
                    continue
                
                conta_str = str(conta).strip()
                nome_str = str(nome).strip() if not pd.isna(nome) else 'N/A'
                ccusto_str = str(ccusto).strip() if not pd.isna(ccusto) else 'N/A'
                
                missing_items.append({
                    'conta': conta_str,
                    'nome': nome_str,
                    'ccusto': ccusto_str
                })
            
            if not missing_items:
                messagebox.showinfo("Info", "Nenhum registro válido encontrado no Excel.")
                return
            
            self.write_log(f"\n{'='*50}")
            self.write_log(f"🔍 BUSCA ASSISTIDA - Excel Completo")
            self.write_log(f"{'='*50}")
            self.write_log(f"📊 Total de registros: {len(missing_items)}")
        
        else:
            # Cancelado
            return
        
        # Abrir janela de busca assistida
        self.open_search_window(missing_items)
    
    def parse_missing_txt(self, txt_path):
        """Lê arquivo TXT e extrai informações dos não encontrados"""
        items = []
        try:
            with open(txt_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            current = None
            for raw in lines:
                line = raw.strip()
                if not line:
                    continue

                # New block starts with pattern like: '1. PDF: filename.pdf'
                m = re.match(r'^\s*\d+\.\s*PDF:\s*(.+)$', line, re.IGNORECASE)
                if m:
                    if current:
                        # Ensure keys exist
                        current.setdefault('conta', 'N/A')
                        current.setdefault('nome', 'N/A')
                        current.setdefault('ccusto', 'N/A')
                        items.append(current)
                    current = {'pdf': m.group(1).strip(), 'conta': 'N/A', 'nome': 'N/A', 'ccusto': 'N/A'}
                    continue

                # If the file was produced by the older format (Conta:, Nome:, Centro de Custo:)
                if line.startswith('Conta:'):
                    if not current:
                        current = {'conta': 'N/A', 'nome': 'N/A', 'ccusto': 'N/A'}
                    current['conta'] = line.split('Conta:', 1)[1].strip()
                    continue
                if line.startswith('Nome:'):
                    if not current:
                        current = {'conta': 'N/A', 'nome': 'N/A', 'ccusto': 'N/A'}
                    current['nome'] = line.split('Nome:', 1)[1].strip()
                    continue
                if line.startswith('Centro de Custo:'):
                    if not current:
                        current = {'conta': 'N/A', 'nome': 'N/A', 'ccusto': 'N/A'}
                    current['ccusto'] = line.split('Centro de Custo:', 1)[1].strip()
                    continue

                # Parse the report format produced by this tool: 'Conta encontrada:' and 'Agência encontrada:'
                if line.lower().startswith('conta encontrada:'):
                    if not current:
                        current = {'conta': 'N/A', 'nome': 'N/A', 'ccusto': 'N/A'}
                    current['conta'] = line.split(':', 1)[1].strip()
                    continue
                if line.lower().startswith('agência encontrada:') or line.lower().startswith('agencia encontrada:'):
                    # We don't use agência here for the assisted search input, but keep it in case
                    if not current:
                        current = {'conta': 'N/A', 'nome': 'N/A', 'ccusto': 'N/A'}
                    # store as agencia (not used for search input)
                    current.setdefault('agencia', line.split(':', 1)[1].strip())
                    continue

                # Also accept lines like 'Página:' (ignored for search but could be stored)
                if line.startswith('Página:') or line.startswith('Pagina:'):
                    if current:
                        try:
                            current['pagina'] = int(line.split(':', 1)[1].strip())
                        except Exception:
                            current['pagina'] = line.split(':', 1)[1].strip()
                    continue

            # Append the last item
            if current:
                current.setdefault('conta', 'N/A')
                current.setdefault('nome', 'N/A')
                current.setdefault('ccusto', 'N/A')
                items.append(current)

        except Exception as e:
            self.write_log(f"❌ Erro ao ler arquivo: {e}")

        # Normalize to expected keys for open_search_window: conta, nome, ccusto
        normalized = []
        for it in items:
            normalized.append({
                'conta': it.get('conta', 'N/A'),
                'nome': it.get('nome', 'N/A'),
                'ccusto': it.get('ccusto', 'N/A')
            })

        return normalized
    
    def open_search_window(self, missing_items):
        """Abre janela interativa para buscar e confirmar comprovantes"""
        ACCENT = self.ACCENT
        search_win = ctk.CTkToplevel(self.root)
        search_win.title("Busca Assistida - Comprovantes Não Encontrados")
        search_win.geometry("1050x720")
        search_win.transient(self.root)

        main_frame = ctk.CTkFrame(search_win, fg_color="transparent")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        ctk.CTkLabel(main_frame, text="Busca Assistida de Comprovantes",
                     font=("Segoe UI", 14, "bold"), text_color=ACCENT).pack(pady=(0, 6))

        info_text = (f"Total de comprovantes não encontrados: {len(missing_items)}\n"
                     "Selecione um item e clique em 'Buscar' para procurar nos PDFs com critérios flexíveis.")
        ctk.CTkLabel(main_frame, text=info_text, font=("Segoe UI", 9)).pack(pady=(0, 10))

        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Lista (esquerda)
        list_frame = ctk.CTkFrame(content_frame, corner_radius=8)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        ctk.CTkLabel(list_frame, text="📋 Não Encontrados",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=10, pady=(8, 4))

        tree_container = ctk.CTkFrame(list_frame, fg_color="transparent")
        tree_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        columns = ('conta', 'nome', 'ccusto')
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=18)
        tree.heading('conta', text='Conta')
        tree.heading('nome', text='Nome')
        tree.heading('ccusto', text='Centro de Custo')
        tree.column('conta', width=100)
        tree.column('nome', width=250)
        tree.column('ccusto', width=150)
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for item in missing_items:
            tree.insert('', tk.END, values=(item.get('conta', ''), item.get('nome', ''), item.get('ccusto', '')))

        # Resultados (direita)
        results_frame = ctk.CTkFrame(content_frame, corner_radius=8)
        results_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        ctk.CTkLabel(results_frame, text="🔍 Resultados da Busca",
                     font=("Segoe UI", 10, "bold"), text_color=ACCENT).pack(anchor=tk.W, padx=10, pady=(8, 4))

        results_text = ctk.CTkTextbox(results_frame, font=("Courier New", 9), state='disabled')
        results_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # Botões
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill=tk.X, pady=(10, 0))

        status_var = tk.StringVar(value="Selecione um item e clique em Buscar")
        status_lbl = ctk.CTkLabel(button_frame, text="Selecione um item e clique em Buscar",
                                   font=("Segoe UI", 9), text_color="gray")
        status_lbl.pack(side=tk.LEFT, padx=(0, 10))
        status_var.trace_add('write', lambda *_: status_lbl.configure(text=status_var.get()))
        
        # Variável para armazenar resultados da busca atual
        current_results = {'matches': [], 'selected_item': None}
        
        def search_selected():
            """Busca o item selecionado nos PDFs (roda em thread para não travar a UI)"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Aviso", "Selecione um item para buscar!")
                return

            item_id = selection[0]
            values = tree.item(item_id)['values']
            conta = values[0]
            nome = values[1]
            ccusto = values[2]

            current_results['selected_item'] = {'conta': conta, 'nome': nome, 'ccusto': ccusto}

            # Preparar UI antes de rodar a busca
            status_var.set(f"Buscando: {nome}...")
            results_text.configure(state='normal')
            results_text.delete("0.0", "end")
            results_text.insert("end", f"Buscando por:\n")
            results_text.insert("end", f"  Conta: {conta}\n")
            results_text.insert("end", f"  Nome: {nome}\n")
            results_text.insert("end", f"  C.Custo: {ccusto}\n")
            results_text.insert("end", f"\n{'='*50}\n\n")
            results_text.configure(state='disabled')

            def worker():
                try:
                    matches = self.flexible_search(conta, nome, ccusto)
                except Exception as e:
                    matches = []
                    err = e
                else:
                    err = None

                def finish_ui():
                    current_results['matches'] = matches
                    results_text.configure(state='normal')
                    results_text.delete("0.0", "end")
                    if err:
                        results_text.insert("end", f"❌ Erro durante a busca: {err}\n")
                        status_var.set("Erro na busca")
                    elif matches:
                        results_text.insert("end", f"✓ Encontrados {len(matches)} possíveis matches:\n\n")
                        for i, match in enumerate(matches, 1):
                            results_text.insert("end", f"{i}. PDF: {match['pdf']}\n")
                            results_text.insert("end", f"   Página: {match['page'] + 1}\n")
                            results_text.insert("end", f"   Critério: {match.get('criteria','?')}\n")
                            results_text.insert("end", f"   Trecho:\n")
                            results_text.insert("end", f"   {match.get('snippet','')}\n")
                            results_text.insert("end", f"\n{'-'*50}\n\n")
                        status_var.set(f"Encontrados {len(matches)} possíveis matches - Revise e confirme")
                    else:
                        results_text.insert("end", "❌ Nenhum match encontrado mesmo com busca flexível.\n\n")
                        results_text.insert("end", "Dicas:\n")
                        results_text.insert("end", "• Verifique se o nome está correto\n")
                        results_text.insert("end", "• Verifique se a conta está correta\n")
                        results_text.insert("end", "• Verifique se o comprovante está no PDF\n")
                        status_var.set("Nenhum match encontrado")
                    results_text.configure(state='disabled')

                # Agendar atualização da UI
                self.root.after(0, finish_ui)

            # Rodar busca em thread separada para não travar a interface
            threading.Thread(target=worker, daemon=True).start()
        
        def extract_selected():
            """Extrai os matches selecionados"""
            if not current_results['matches']:
                messagebox.showwarning("Aviso", "Faça uma busca primeiro!")
                return
            
            # Abrir diálogo de confirmação com lista de matches
            confirm_msg = f"Confirmar extração de {len(current_results['matches'])} comprovante(s)?\n\n"
            for match in current_results['matches']:
                confirm_msg += f"• {match['pdf']} - Pág {match['page'] + 1}\n"
            
            if not messagebox.askyesno("Confirmar Extração", confirm_msg):
                return
            
            # Extrair
            item = current_results['selected_item']
            out_dir = normalize_path(self.out_var.get() or "comprovantes_extraidos")
            pdf_folder = normalize_path(self.pdf_folder_var.get())
            
            success_count = 0
            for match in current_results['matches']:
                pdf_path = os.path.join(pdf_folder, match['pdf'])
                nome_str = clean_filename(item['nome'])
                ccusto_str = clean_filename(item['ccusto'])
                
                # Criar subpasta para o centro de custo
                ccusto_folder = os.path.join(out_dir, ccusto_str)
                Path(ccusto_folder).mkdir(parents=True, exist_ok=True)
                
                # Salvar na pasta do ccusto (mantém prefixo de ccusto, com sufixo _manual)
                out_path = os.path.join(ccusto_folder, f"{ccusto_str}_{nome_str}_manual.pdf")
                i = 1
                while os.path.exists(out_path):
                    out_path = os.path.join(ccusto_folder, f"{ccusto_str}_{nome_str}_manual_{i}.pdf")
                    i += 1
                
                pages_written = create_pdf(pdf_path, [match['page']], out_path)
                if pages_written and pages_written > 0:
                    # Somar pelo número de páginas extraídas (normalmente 1 neste fluxo manual)
                    success_count += pages_written
                    self.write_log(f"✓ Extraído manualmente: {ccusto_str}/{ccusto_str}_{nome_str}_manual (pág {match['page'] + 1})")
            
            messagebox.showinfo("Sucesso", f"{success_count} comprovante(s) extraído(s) com sucesso!")
            status_var.set(f"Extraídos {success_count} comprovantes")
            
            # Remover item da lista
            if success_count > 0:
                tree.delete(tree.selection())
        
        ctk.CTkButton(button_frame, text="🔍 Buscar", command=search_selected, width=120,
                      fg_color=ACCENT, hover_color=self.ACCENT_HOVER).pack(side=tk.RIGHT, padx=(5, 0))
        ctk.CTkButton(button_frame, text="✓ Extrair Selecionados", command=extract_selected, width=180,
                      fg_color="#43A047", hover_color="#2E7D32").pack(side=tk.RIGHT, padx=(5, 0))
        ctk.CTkButton(button_frame, text="❌ Fechar", command=search_win.destroy, width=100,
                      fg_color="gray40", hover_color="gray30").pack(side=tk.RIGHT)
    
    def flexible_search(self, conta, nome, ccusto):
        """Busca flexível nos PDFs com múltiplos critérios relaxados"""
        matches = []
        pdf_folder = normalize_path(self.pdf_folder_var.get())
        
        # Listar PDFs
        pdf_files = []
        try:
            pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
        except Exception:
            return matches
        
        # Normalizar termos de busca
        def normalize_search_text(s):
            if not s:
                return ""
            nf = unicodedata.normalize('NFKD', str(s))
            ascii_s = nf.encode('ascii', 'ignore').decode('ascii')
            cleaned = re.sub(r'[^A-Za-z0-9\s]', ' ', ascii_s)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip().upper()
            return cleaned
        
        conta_norm = normalize_account(conta)
        nome_norm = normalize_search_text(nome)
        nome_parts = [p for p in nome_norm.split() if len(p) >= 3]
        
        # Buscar em cada PDF
        for pdf_name in pdf_files:
            pdf_path = os.path.join(pdf_folder, pdf_name)
            
            try:
                pages = extract_pdf_pages(pdf_path)
                
                for page_num, page_data in pages.items():
                    text = page_data['text']
                    text_norm = page_data['norm_text']
                    text_numbers = page_data['numbers']
                    
                    criteria_met = []
                    
                    # Critério 1: Conta encontrada
                    if conta_norm and conta_norm in text_numbers:
                        criteria_met.append("Conta exata")
                    
                    # Critério 2: Nome completo encontrado
                    if nome_norm and nome_norm in text_norm:
                        criteria_met.append("Nome completo")
                    
                    # Critério 3: Múltiplas partes do nome (flexível)
                    if nome_parts:
                        found_parts = sum(1 for part in nome_parts if part in text_norm)
                        if found_parts >= max(2, len(nome_parts) // 2):
                            criteria_met.append(f"{found_parts}/{len(nome_parts)} partes do nome")
                    
                    # Critério 4: Primeiro e último nome
                    if len(nome_parts) >= 2:
                        if nome_parts[0] in text_norm and nome_parts[-1] in text_norm:
                            criteria_met.append("Primeiro + último nome")
                    
                    # Se encontrou pelo menos 1 critério, adicionar como candidato
                    if criteria_met:
                        # Extrair snippet (contexto)
                        snippet = self.extract_snippet(text, nome, conta)
                        
                        matches.append({
                            'pdf': pdf_name,
                            'page': page_num,
                            'criteria': ", ".join(criteria_met),
                            'snippet': snippet,
                            'score': len(criteria_met)
                        })
            
            except Exception as e:
                self.write_log(f"⚠️ Erro ao processar {pdf_name}: {e}")
                continue
        
        # Ordenar por score (mais critérios primeiro)
        matches.sort(key=lambda x: x['score'], reverse=True)
        
        return matches
    
    def extract_snippet(self, text, nome, conta, context_chars=150):
        """Extrai trecho do texto ao redor do nome/conta encontrado"""
        text = text or ""
        
        # Tentar encontrar posição do nome
        nome_clean = str(nome).strip()
        pos = text.upper().find(nome_clean.upper())
        
        if pos == -1:
            # Tentar conta
            conta_clean = str(conta).strip()
            pos = text.find(conta_clean)
        
        if pos == -1:
            # Retornar início do texto
            snippet = text[:context_chars * 2]
        else:
            # Extrair contexto ao redor
            start = max(0, pos - context_chars)
            end = min(len(text), pos + len(nome_clean) + context_chars)
            snippet = text[start:end]
        
        # Limpar e formatar
        snippet = ' '.join(snippet.split())
        if len(snippet) > 300:
            snippet = snippet[:300] + "..."
        
        return snippet
    
    def diagnose_missing(self, conta_info, pdf_files, pdf_folder):
        """Diagnostica por que um comprovante não foi encontrado"""
        conta = conta_info['conta']
        nome = conta_info['nome']
        
        # Normalizar para busca
        def normalize_search_text(s):
            if not s:
                return ""
            nf = unicodedata.normalize('NFKD', str(s))
            ascii_s = nf.encode('ascii', 'ignore').decode('ascii')
            cleaned = re.sub(r'[^A-Za-z0-9\s]', ' ', ascii_s)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip().upper()
            return cleaned
        
        conta_norm = normalize_account(conta)
        nome_norm = normalize_search_text(nome)
        nome_parts = [p for p in nome_norm.split() if len(p) >= 3]
        
        pdfs_com_conta = []
        pdfs_com_nome = []
        pdfs_com_ambos_separados = []
        
        # Cache de páginas extraídas para evitar reprocessamento
        if not hasattr(self, '_pdf_cache'):
            self._pdf_cache = {}
        
        # Verificar cada PDF
        for pdf_name in pdf_files:
            pdf_path = os.path.join(pdf_folder, pdf_name)
            
            try:
                # Usar cache se disponível
                if pdf_path not in self._pdf_cache:
                    self._pdf_cache[pdf_path] = extract_pdf_pages(pdf_path)
                
                pages = self._pdf_cache[pdf_path]
                
                tem_conta_pdf = False
                tem_nome_pdf = False
                paginas_com_conta = []
                paginas_com_nome = []
                
                for page_num, page_data in pages.items():
                    text_norm = page_data['norm_text']
                    text_numbers = page_data['numbers']
                    
                    # Verificar conta
                    if conta_norm and conta_norm in text_numbers:
                        tem_conta_pdf = True
                        paginas_com_conta.append(page_num + 1)
                    
                    # Verificar nome
                    if nome_norm and nome_norm in text_norm:
                        tem_nome_pdf = True
                        paginas_com_nome.append(page_num + 1)
                    else:
                        # Verificar partes do nome
                        if nome_parts:
                            found_parts = sum(1 for part in nome_parts if part in text_norm)
                            if found_parts >= max(2, len(nome_parts) // 2):
                                tem_nome_pdf = True
                                paginas_com_nome.append(page_num + 1)
                
                if tem_conta_pdf:
                    pdfs_com_conta.append(f"{pdf_name} (pág {paginas_com_conta})")
                
                if tem_nome_pdf:
                    pdfs_com_nome.append(f"{pdf_name} (pág {paginas_com_nome})")
                
                # Verificar se tem ambos mas em páginas diferentes
                if tem_conta_pdf and tem_nome_pdf:
                    # Ver se há intersecção de páginas
                    if not set(paginas_com_conta).intersection(set(paginas_com_nome)):
                        pdfs_com_ambos_separados.append(pdf_name)
                
            except Exception:
                continue
        
        # Montar diagnóstico
        diagnostico = {
            'encontrou_conta': len(pdfs_com_conta) > 0,
            'encontrou_nome': len(pdfs_com_nome) > 0,
            'pdfs_com_conta': pdfs_com_conta[:3],  # Limitar a 3 para não poluir
            'pdfs_com_nome': pdfs_com_nome[:3],
            'tipo': '',
            'detalhes': '',
            'sugestoes': []
        }
        
        # Determinar tipo de problema
        if not diagnostico['encontrou_conta'] and not diagnostico['encontrou_nome']:
            diagnostico['tipo'] = 'Conta e Nome não encontrados'
            diagnostico['detalhes'] = 'Nenhum dos dados (conta ou nome) foi encontrado em nenhum PDF'
            diagnostico['sugestoes'] = [
                'Verifique se a conta e o nome estão corretos no Excel',
                'Confirme se o comprovante desta pessoa está nos PDFs fornecidos',
                'Verifique se há erros de digitação nos dados'
            ]
        
        elif diagnostico['encontrou_conta'] and not diagnostico['encontrou_nome']:
            diagnostico['tipo'] = 'Conta encontrada, Nome não'
            diagnostico['detalhes'] = f'A conta foi encontrada, mas o nome "{nome}" não aparece nas mesmas páginas'
            diagnostico['sugestoes'] = [
                'O nome no Excel pode estar diferente do nome no PDF',
                'Verifique variações do nome (abreviações, nome completo vs nome social)',
                'Use a busca assistida para ver o que está na página com esta conta'
            ]
        
        elif not diagnostico['encontrou_conta'] and diagnostico['encontrou_nome']:
            diagnostico['tipo'] = 'Nome encontrado, Conta não'
            diagnostico['detalhes'] = f'O nome foi encontrado, mas a conta "{conta}" não aparece nas mesmas páginas'
            diagnostico['sugestoes'] = [
                'A conta no Excel pode estar incorreta ou diferente do PDF',
                'Verifique se a conta tem dígito verificador ou formatação diferente',
                'Use a busca assistida para ver qual conta está associada a este nome'
            ]
        
        elif pdfs_com_ambos_separados:
            diagnostico['tipo'] = 'Ambos em PDFs diferentes'
            diagnostico['detalhes'] = 'Conta e nome foram encontrados, mas sempre em páginas diferentes do PDF'
            diagnostico['sugestoes'] = [
                'Pode haver homonímia (duas pessoas com nomes similares)',
                'A conta pode pertencer a outra pessoa com nome parecido',
                'Verifique manualmente os PDFs listados acima'
            ]
        
        else:
            diagnostico['tipo'] = 'Critérios não atendidos'
            diagnostico['detalhes'] = 'Conta e/ou nome encontrados mas não na mesma página com critérios exigidos'
            diagnostico['sugestoes'] = [
                'Use a busca assistida com critérios flexíveis',
                'Verifique se o formato dos dados no PDF é diferente do esperado'
            ]
        
        return diagnostico
    
    # ==================== GOOGLE DRIVE UPLOAD ====================
    
    def calculate_folder_summary(self, folder_path):
        """Calcula resumo dos arquivos em uma pasta"""
        summary = {
            'total_files': 0,
            'total_folders': 0,
            'total_size': 0,
            'folders': {}
        }
        
        try:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.endswith('.pdf'):
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        
                        # Pegar nome da subpasta (centro de custo)
                        rel_path = os.path.relpath(root, folder_path)
                        if rel_path == '.':
                            ccusto = "Raiz"
                        else:
                            ccusto = rel_path
                        
                        if ccusto not in summary['folders']:
                            summary['folders'][ccusto] = {
                                'count': 0,
                                'size': 0,
                                'files': []
                            }
                        
                        summary['folders'][ccusto]['count'] += 1
                        summary['folders'][ccusto]['size'] += file_size
                        summary['folders'][ccusto]['files'].append(file)
                        
                        summary['total_files'] += 1
                        summary['total_size'] += file_size
            
            summary['total_folders'] = len(summary['folders'])
            
        except Exception as e:
            self.write_log(f"⚠️ Erro ao calcular resumo: {e}")
        
        return summary
    
    def format_size(self, size_bytes):
        """Formata tamanho em bytes para formato legível"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"
    
    def detect_google_drive_folder(self):
        """Tenta detectar pasta do Google Drive automaticamente"""
        username = os.getlogin() if platform.system() == 'Windows' else os.path.expanduser("~").split('/')[-1]
        
        possible_paths = [
            os.path.expanduser("~/Google Drive"),
            os.path.expanduser("~/GoogleDrive"),
            f"C:/Users/{username}/Google Drive",
            f"C:/Users/{username}/GoogleDrive",
            os.path.expanduser("~/Google Drive/My Drive"),
            os.path.expanduser("~/OneDrive"),
        ]
        
        for path in possible_paths:
            if os.path.exists(path) and os.path.isdir(path):
                return path
        
        return None
    
    def open_drive_upload_dialog(self):
        """Abre janela para revisar e enviar arquivos para Google Drive"""
        if not self.last_output_folder or not os.path.exists(self.last_output_folder):
            messagebox.showwarning("Aviso", "Nenhuma pasta de saída encontrada. Execute o processamento primeiro.")
            return
        
        # Calcular resumo dos arquivos
        summary = self.calculate_folder_summary(self.last_output_folder)
        
        if summary['total_files'] == 0:
            messagebox.showinfo("Info", "Nenhum arquivo PDF encontrado na pasta de saída.")
            return
        
        # Criar janela de diálogo
        DriveUploadDialog(self.root, self, self.last_output_folder, summary)
    
    def upload_to_drive(self, source_folder, drive_destination, options):
        """Faz upload dos arquivos para Google Drive (cópia para pasta sincronizada)"""
        
        # Coletar todos os arquivos
        files_to_upload = []
        total_size = 0
        
        try:
            for root, dirs, files in os.walk(source_folder):
                for file in files:
                    if file.endswith('.pdf'):
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        
                        # Determinar pasta destino (manter estrutura de centro de custo)
                        rel_path = os.path.relpath(root, source_folder)
                        
                        files_to_upload.append({
                            'source': file_path,
                            'destination': os.path.join(drive_destination, rel_path, file) if rel_path != '.' else os.path.join(drive_destination, file),
                            'size': file_size,
                            'ccusto': rel_path if rel_path != '.' else 'Raiz'
                        })
                        
                        total_size += file_size
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao listar arquivos: {e}")
            return
        
        if not files_to_upload:
            messagebox.showinfo("Info", "Nenhum arquivo PDF encontrado para enviar.")
            return
        
        # Criar janela de progresso
        progress_dialog = UploadProgressDialog(self.root, self)
        
        # Variáveis de controle
        uploaded = 0
        errors = []
        bytes_sent = 0
        start_time = time.time()
        
        # Thread de upload
        def upload_worker():
            nonlocal uploaded, bytes_sent
            
            for i, file_info in enumerate(files_to_upload, 1):
                # Verificar se foi cancelado
                if progress_dialog.cancelled:
                    break
                
                try:
                    # Criar diretório destino se não existir
                    dest_dir = os.path.dirname(file_info['destination'])
                    os.makedirs(dest_dir, exist_ok=True)
                    
                    # Verificar se arquivo já existe
                    if os.path.exists(file_info['destination']):
                        # Comparar tamanhos
                        src_size = os.path.getsize(file_info['source'])
                        dst_size = os.path.getsize(file_info['destination'])
                        
                        if src_size == dst_size:
                            # Arquivo idêntico, pular
                            self.write_log(f"⏭️ Pulado (já existe): {os.path.basename(file_info['source'])}")
                            uploaded += 1
                            bytes_sent += file_info['size']
                            continue
                        else:
                            # Arquivo diferente, criar nome alternativo
                            base, ext = os.path.splitext(file_info['destination'])
                            counter = 1
                            while os.path.exists(file_info['destination']):
                                file_info['destination'] = f"{base}_{counter}{ext}"
                                counter += 1
                    
                    # Copiar arquivo
                    shutil.copy2(file_info['source'], file_info['destination'])
                    
                    uploaded += 1
                    bytes_sent += file_info['size']
                    
                    # Atualizar progresso
                    elapsed = time.time() - start_time
                    can_continue = progress_dialog.update_progress(
                        current=i,
                        total=len(files_to_upload),
                        current_file=file_info['source'],
                        bytes_sent=bytes_sent,
                        bytes_total=total_size,
                        elapsed_time=elapsed
                    )
                    
                    if not can_continue:
                        break
                    
                except Exception as e:
                    errors.append({
                        'file': file_info['source'],
                        'error': str(e)
                    })
                    self.write_log(f"❌ Erro ao copiar {os.path.basename(file_info['source'])}: {e}")
            
            # Finalizar
            duration = time.time() - start_time
            
            # Fechar janela de progresso
            self.root.after(0, lambda: progress_dialog.close())
            
            # Remover pasta local se solicitado
            if options.get('keep_local') == False and not progress_dialog.cancelled:
                try:
                    shutil.rmtree(source_folder)
                    self.write_log(f"🗑️ Pasta local removida: {source_folder}")
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao remover pasta local: {e}")
            
            # Mostrar resultado
            results = {
                'success': uploaded,
                'errors': len(errors),
                'error_list': errors,
                'duration': str(timedelta(seconds=int(duration))),
                'size_mb': round(total_size / (1024 * 1024), 2),
                'drive_url': drive_destination,
                'cancelled': progress_dialog.cancelled
            }
            
            # Abrir pasta do Drive se solicitado
            if options.get('open_after') and not progress_dialog.cancelled and uploaded > 0:
                try:
                    if platform.system() == 'Windows':
                        os.startfile(drive_destination)
                    elif platform.system() == 'Darwin':
                        subprocess.Popen(['open', drive_destination])
                    else:
                        subprocess.Popen(['xdg-open', drive_destination])
                except:
                    pass
            
            # Mostrar relatório final
            if not progress_dialog.cancelled:
                self.root.after(0, lambda: UploadCompleteDialog(self.root, self, results))
                
                # Log resumo
                self.write_log(f"\n{'='*50}")
                self.write_log(f"📤 UPLOAD CONCLUÍDO")
                self.write_log(f"{'='*50}")
                self.write_log(f"✓ Enviados: {uploaded}")
                self.write_log(f"✗ Erros: {len(errors)}")
                self.write_log(f"⏱️ Tempo: {results['duration']}")
                self.write_log(f"💾 Tamanho: {results['size_mb']} MB")
                self.write_log(f"🔗 Destino: {drive_destination}")
                self.write_log(f"{'='*50}")
            else:
                self.write_log(f"\n⚠️ Upload cancelado pelo usuário")
                messagebox.showinfo("Cancelado", 
                    f"Upload cancelado.\n\n{uploaded} de {len(files_to_upload)} arquivo(s) foram enviados antes do cancelamento.")
        
        # Iniciar thread de upload
        threading.Thread(target=upload_worker, daemon=True).start()
    
    def start(self):
        if not self.pdf_folder_var.get() or not self.excel_var.get():
            messagebox.showerror("Erro", "Selecione a pasta de PDFs e o Excel!")
            return
        if self.df is None:
            messagebox.showerror("Erro", "Carregue Excel!")
            return
        if not self.conta_col or not self.agencia_col or not self.nome_col or not self.ccusto_col:
            messagebox.showerror("Erro", "Colunas não encontradas no Excel!\nVerifique se existem as colunas: Conta, Agência e Nome\n(o Centro de Custo é detectado automaticamente pelo nome da aba)")
            return

        self.btn.configure(state='disabled')
        self.status_var.set("Processando...")
        self.prog.start()
        self.start_timer()
        threading.Thread(target=self.process, daemon=True).start()
    
    def process(self):
        try:
            pdf_folder = normalize_path(self.pdf_folder_var.get())
            out_dir = normalize_path(self.out_var.get())
            conta_col = self.conta_col
            agencia_col = self.agencia_col
            nome_col = self.nome_col
            ccusto_col = self.ccusto_col
            
            # Verificar se as pastas existem
            if not os.path.exists(pdf_folder) or not os.path.isdir(pdf_folder):
                self.write_log(f"❌ Pasta de PDFs não encontrada: {pdf_folder}")
                messagebox.showerror("Erro", f"Pasta de PDFs não encontrada")
                return
            
            Path(out_dir).mkdir(parents=True, exist_ok=True)
            
            self.write_log("\n" + "="*50)
            self.write_log("🚀 Iniciando processamento...")
            self.write_log("="*50)
            
            # Listar todos os PDFs na pasta usando múltiplos métodos (compatível com OneDrive)
            pdf_files_set = set()
            
            # Método 1: os.listdir
            try:
                files_listdir = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
                pdf_files_set.update(files_listdir)
                self.write_log(f"ℹ️ Método listdir: {len(files_listdir)} PDFs")
            except Exception as e:
                self.write_log(f"⚠️ Erro com listdir: {e}")
            
            # Método 2: Path.iterdir (confiável para OneDrive)
            try:
                path_obj = Path(pdf_folder)
                files_iterdir = [f.name for f in path_obj.iterdir() if f.is_file() and f.suffix.lower() == '.pdf']
                pdf_files_set.update(files_iterdir)
                self.write_log(f"ℹ️ Método iterdir: {len(files_iterdir)} PDFs")
            except Exception as e:
                self.write_log(f"⚠️ Erro com iterdir: {e}")
            
            # Método 3: os.scandir (eficiente)
            try:
                with os.scandir(pdf_folder) as entries:
                    files_scandir = [e.name for e in entries if e.is_file() and e.name.lower().endswith('.pdf')]
                pdf_files_set.update(files_scandir)
                self.write_log(f"ℹ️ Método scandir: {len(files_scandir)} PDFs")
            except Exception as e:
                self.write_log(f"⚠️ Erro com scandir: {e}")
            
            pdf_files = sorted(list(pdf_files_set))
            
            if not pdf_files:
                self.write_log("\n⚠️ Nenhum PDF encontrado na pasta!")
                self.write_log("   💡 Dica: Se os arquivos estão no OneDrive, tente:")
                self.write_log("      1. Verificar se os PDFs foram baixados localmente")
                self.write_log("      2. Clicar com botão direito nos PDFs > 'Sempre manter neste dispositivo'")
                self.write_log("      3. Ou mover os PDFs para uma pasta local fora do OneDrive")
                return
            
            self.write_log(f"\n📊 Total de PDFs encontrados: {len(pdf_files)}")
            
            # Separar PDFs novos e já processados (ou forçar reprocessamento)
            novos_pdfs = []
            ja_processados = []
            force = getattr(self, 'force_reprocess_var', None) and self.force_reprocess_var.get()
            if force:
                self.write_log("⚠️ Modo FORÇAR reprocessamento ativo: ignorando histórico e reprocessando todos os PDFs.")

            for pdf_name in pdf_files:
                pdf_path = os.path.join(pdf_folder, pdf_name)
                fingerprint = self.get_pdf_fingerprint(pdf_path)

                if (not force) and fingerprint and fingerprint in self.processed_pdfs:
                    ja_processados.append(pdf_name)
                else:
                    novos_pdfs.append((pdf_name, pdf_path, fingerprint))
            
            if ja_processados:
                self.write_log(f"⏭️ PDFs já processados anteriormente: {len(ja_processados)}")
            
            if not novos_pdfs:
                self.write_log("\n✓ Todos os PDFs já foram processados!")
                elapsed = self.stop_timer()
                time_str = self.format_time(elapsed)
                self.write_log(f"⏱️ Tempo total: {time_str}")
                self.root.after(0, lambda: self.status_var.set("Concluído - Nenhum PDF novo"))
                self.root.after(0, lambda: messagebox.showinfo(
                    "Processamento Concluído", 
                    f"Todos os {len(pdf_files)} PDFs já foram processados anteriormente!"
                ))
                return
            
            self.write_log(f"🆕 PDFs novos para processar: {len(novos_pdfs)}")
            self.root.after(0, lambda: self.status_var.set(f"Processando {len(novos_pdfs)} PDFs..."))
            
            # Processamento dos PDFs novos
            total_ok = 0
            total_nok = 0
            total_duplicates = 0
            
            # Dicionário para rastrear quais contas foram encontradas
            contas_encontradas = set()  # Conjunto de contas que foram extraídas com sucesso
            todas_contas = []  # Lista de todas as contas do Excel para verificar no final
            
            # Primeiro, coletar todas as contas do Excel
            for row_idx, row in self.df.iterrows():
                conta = row[conta_col]
                agencia = row[agencia_col]
                nome = row[nome_col]
                ccusto = row[ccusto_col]
                
                # Campos obrigatórios
                if pd.isna(nome) or str(nome).strip() == '':
                    continue
                if pd.isna(ccusto) or str(ccusto).strip() == '':
                    continue
                
                # Para conta e agência, buscar em TODAS as colunas se estiverem vazias
                conta_str = str(conta).strip() if not pd.isna(conta) and str(conta).strip() != '' else None
                agencia_str = str(agencia).strip() if not pd.isna(agencia) and str(agencia).strip() != '' else None
                
                # Se conta ou agência estão vazias, procurar em OUTRAS COLUNAS
                valores_encontrados = []
                if not conta_str or not agencia_str:
                    # Percorrer todas as colunas buscando valores numéricos
                    for col_name in row.index:
                        if col_name in [nome_col, ccusto_col]:  # Pular colunas de texto
                            continue
                        
                        valor = row[col_name]
                        if pd.isna(valor):
                            continue
                        
                        valor_str = str(valor).strip()
                        # Verificar se é um valor numérico válido (pode ter hífen para DV)
                        if valor_str and re.match(r'^[\d\-\.]+$', valor_str):
                            valor_norm = normalize_account(valor_str)
                            if valor_norm and len(valor_norm) >= 3:
                                valores_encontrados.append(valor_str)
                    
                    # Se encontrou valores, usar os primeiros 2
                    if len(valores_encontrados) >= 2:
                        if not conta_str:
                            conta_str = valores_encontrados[0]
                        if not agencia_str:
                            agencia_str = valores_encontrados[1] if len(valores_encontrados) > 1 else valores_encontrados[0]
                    elif len(valores_encontrados) == 1:
                        # Só tem 1 valor, usar como conta
                        if not conta_str:
                            conta_str = valores_encontrados[0]
                        if not agencia_str:
                            # Tentar usar o mesmo valor como agência (pode estar duplicado)
                            agencia_str = valores_encontrados[0]
                
                # Se ainda não tem conta E agência, pular este registro
                if not conta_str or not agencia_str:
                    continue
                
                nome_str = str(nome).strip() if not pd.isna(nome) else 'N/A'
                ccusto_str = str(ccusto).strip() if not pd.isna(ccusto) else 'N/A'
                
                todas_contas.append({
                    'conta': conta_str,
                    'agencia': agencia_str,
                    'nome': nome_str,
                    'ccusto': ccusto_str
                })
            
            # Rastrear páginas processadas
            total_paginas_pdfs = 0
            paginas_com_match = set()  # páginas que tiveram match (PDF + número da página)
            paginas_ja_extraidas = set()  # Controle de páginas já extraídas (evita duplicatas)
            autenticacoes_ja_extraidas = set()  # Controle por código de autenticação (deduplicação robusta)
            
            for idx, (pdf_name, pdf_path, fingerprint) in enumerate(novos_pdfs, 1):
                self.write_log(f"\n{'='*50}")
                self.write_log(f"📄 Processando PDF {idx}/{len(novos_pdfs)}: {pdf_name}")
                self.write_log(f"{'='*50}")
                self.root.after(0, lambda i=idx, t=len(novos_pdfs): self.status_var.set(f"PDF {i}/{t}..."))
                
                try:
                    pages = extract_pdf_pages(pdf_path)
                    total_paginas_pdfs += len(pages)
                    self.write_log(f"📄 Total de páginas neste PDF: {len(pages)}")
                    
                    ok = 0
                    nok = 0
                    duplicates = 0
                    
                    for row_idx, row in self.df.iterrows():
                        conta = row[conta_col]
                        agencia = row[agencia_col]
                        nome = row[nome_col]
                        ccusto = row[ccusto_col]
                        
                        # Verificar campos obrigatórios (nome e ccusto são essenciais)
                        if pd.isna(nome) or str(nome).strip() == '':
                            continue
                        if pd.isna(ccusto) or str(ccusto).strip() == '':
                            continue
                        
                        # Para conta e agência, buscar em TODAS as colunas se estiverem vazias
                        conta_str = str(conta).strip() if not pd.isna(conta) and str(conta).strip() != '' else None
                        agencia_str = str(agencia).strip() if not pd.isna(agencia) and str(agencia).strip() != '' else None
                        
                        # Se conta ou agência estão vazias, procurar em OUTRAS COLUNAS
                        valores_encontrados = []
                        busca_alternativa = False
                        if not conta_str or not agencia_str:
                            busca_alternativa = True
                            # Percorrer todas as colunas buscando valores numéricos
                            for col_name in row.index:
                                if col_name in [nome_col, ccusto_col]:  # Pular colunas de texto
                                    continue
                                
                                valor = row[col_name]
                                if pd.isna(valor):
                                    continue
                                
                                valor_str = str(valor).strip()
                                # Verificar se é um valor numérico válido (pode ter hífen para DV)
                                if valor_str and re.match(r'^[\d\-\.]+$', valor_str):
                                    valor_norm = normalize_account(valor_str)
                                    if valor_norm and len(valor_norm) >= 3:
                                        valores_encontrados.append(valor_str)
                            
                            # Se encontrou valores, usar os primeiros 2
                            if len(valores_encontrados) >= 2:
                                if not conta_str:
                                    conta_str = valores_encontrados[0]
                                if not agencia_str:
                                    agencia_str = valores_encontrados[1] if len(valores_encontrados) > 1 else valores_encontrados[0]
                            elif len(valores_encontrados) == 1:
                                # Só tem 1 valor, usar como conta
                                if not conta_str:
                                    conta_str = valores_encontrados[0]
                                if not agencia_str:
                                    # Tentar usar o mesmo valor como agência (pode estar duplicado)
                                    agencia_str = valores_encontrados[0]
                        
                        # Se ainda não tem conta E agência, pular
                        if not conta_str or not agencia_str:
                            continue
                        
                        nome_str = clean_filename(str(nome).strip())
                        ccusto_str = clean_filename(str(ccusto).strip())
                        
                        # Log se usou busca alternativa
                        if busca_alternativa and valores_encontrados:
                            if self.debug_mode_var.get():
                                self.write_log(f"  📌 {nome_str}: Valores encontrados em colunas alternativas (Conta={conta_str}, Ag={agencia_str})")
                        
                        paginas, valores_invertidos = find_account_pages(conta_str, agencia_str, pages)

                        if paginas:
                            # Filtrar apenas páginas que ainda NÃO foram extraídas
                            # Usa código de autenticação como chave primária (mais confiável que número de página)
                            paginas_novas = []
                            for pag in paginas:
                                chave_pagina = f"{pdf_name}|{pag}"
                                auth = pages[pag].get('auth_code')
                                # Já extraído por autenticação?
                                if auth and auth in autenticacoes_ja_extraidas:
                                    continue
                                # Já extraído por página (fallback para páginas sem auth code)?
                                if chave_pagina in paginas_ja_extraidas:
                                    continue
                                paginas_novas.append(pag)

                            # Se não há páginas novas, pular
                            if not paginas_novas:
                                continue

                            # Criar subpasta para o centro de custo
                            ccusto_folder = os.path.join(out_dir, ccusto_str)
                            Path(ccusto_folder).mkdir(parents=True, exist_ok=True)

                            # Tentar extrair o nome diretamente do conteúdo do PDF (primeira página com match)
                            nome_no_pdf = None
                            for pag in paginas_novas:
                                nome_no_pdf = extract_name_from_page(pages[pag])
                                if nome_no_pdf:
                                    break

                            # Usar nome do PDF se encontrado; caso contrário, usar nome da planilha
                            if nome_no_pdf:
                                nome_final = clean_filename(nome_no_pdf)
                                if self.debug_mode_var.get():
                                    self.write_log(f"  📝 Nome extraído do PDF: '{nome_no_pdf}' (planilha: '{nome_str}')")
                            else:
                                nome_final = nome_str

                            # Salvar PDF na pasta do centro de custo (mantém prefixo de ccusto no nome)
                            out = os.path.join(ccusto_folder, f"{ccusto_str}_{nome_final}.pdf")
                            i = 1
                            while os.path.exists(out):
                                out = os.path.join(ccusto_folder, f"{ccusto_str}_{nome_final}_{i}.pdf")
                                i += 1

                            # Tentar criar o PDF com as páginas novas e obter quantas páginas foram gravadas
                            pages_written = create_pdf(pdf_path, paginas_novas, out)
                            if pages_written and pages_written > 0:
                                # Registrar quais páginas tiveram match (apenas após gravação bem-sucedida)
                                for pag in paginas_novas:
                                    paginas_com_match.add(f"{pdf_name}|{pag}")
                                    paginas_ja_extraidas.add(f"{pdf_name}|{pag}")
                                    auth = pages[pag].get('auth_code')
                                    if auth:
                                        autenticacoes_ja_extraidas.add(auth)

                                self.write_log(f"✓ {ccusto_str}/{ccusto_str}_{nome_final} (pág {[p+1 for p in paginas_novas]})")
                                # Incrementar por número de páginas efetivamente escritas
                                ok += int(pages_written)
                                # Marcar que esta conta foi encontrada
                                contas_encontradas.add(conta_str)
                            else:
                                nok += 1
                    
                    # Registrar PDF como processado
                    if fingerprint:
                        self.processed_pdfs[fingerprint] = {
                            'nome': pdf_name,
                            'data': time.strftime('%d/%m/%Y %H:%M:%S'),
                            'extraidos': ok,
                            'nao_encontrados': nok,
                        }
                        self.save_processed_pdfs()
                    
                    total_ok += ok
                    total_nok += nok
                    total_duplicates += duplicates
                    
                    self.write_log(f"✓ Comprovantes extraídos deste PDF: {ok}")
                    
                except Exception as e:
                    self.write_log(f"❌ Erro ao processar {pdf_name}: {e}")
            
            # Calcular quantas páginas dos PDFs ficaram SEM match com a planilha
            paginas_sem_match = total_paginas_pdfs - len(paginas_com_match)
            
            # Parar timer e calcular tempo total
            elapsed = self.stop_timer()
            time_str = self.format_time(elapsed)
            
            # Comprovantes nos PDFs que NÃO têm funcionário correspondente na planilha
            nao_encontrados = []
            
            # Criar índice de contas+agência do Excel para busca rápida
            # Chave: "conta_agencia" normalizada
            # Também criar índice INVERTIDO para detectar inversões
            contas_excel_set = set()
            contas_excel_invertido_set = set()  # Para detectar inversões
            contas_excel_conta_set = set()  # Índice apenas de contas (conta isolada)
            for conta_info in todas_contas:
                conta_norm = normalize_account(conta_info['conta'])
                agencia_norm = normalize_account(conta_info['agencia'])
                # Indexar conta isolada para permitir match apenas por conta
                if conta_norm:
                    contas_excel_conta_set.add(conta_norm)
                if conta_norm and agencia_norm:
                    # Usar combinação conta+agência como chave única
                    contas_excel_set.add(f"{conta_norm}_{agencia_norm}")
                    # Também adicionar versão invertida para detectar inversões na planilha
                    contas_excel_invertido_set.add(f"{agencia_norm}_{conta_norm}")
            
            self.write_log(f"\n🔍 Analisando páginas sem match para identificar contas não cadastradas...")
            
            # Percorrer todos os PDFs e analisar CADA PÁGINA que não teve match
            for pdf_name in pdf_files:
                pdf_path = os.path.join(pdf_folder, pdf_name)
                try:
                    pages = extract_pdf_pages(pdf_path)
                    
                    for page_num, page_data in pages.items():
                        # Verificar se esta página teve match
                        pagina_id = f"{pdf_name}|{page_num}"
                        if pagina_id in paginas_com_match:
                            continue  # Já foi extraída, pular
                        
                        # BUSCAR APENAS NA SEÇÃO "DADOS DA CONTA CREDITADA"
                        credited_section = page_data.get('credited_section', '')
                        
                        # Se não encontrou a seção, pular esta página
                        if not credited_section or len(credited_section) < 20:
                            continue
                        
                        # Buscar especificamente o campo "Conta corrente:" seguido do número
                        # Padrões possíveis: "Conta corrente: 94894 - 2", "Conta: 12345-6", "C/C: 12345-6"
                        conta_patterns = [
                            r'[Cc]onta\s*[Cc]orrente[:\s]+(\d{4,7}[\s\-]*\d?)',  # Conta corrente: 94894 - 2
                            r'[Cc]/[Cc][:\s]+(\d{4,7}[\s\-]*\d?)',               # C/C: 12345-6
                            r'[Cc]onta[:\s]+(\d{4,7}[\s\-]*\d?)',                # Conta: 12345-6
                        ]
                        
                        # Buscar agência também
                        agencia_patterns = [
                            r'[Aa]g[eê]ncia[:\s]+(\d{3,5})',  # Agência: 6677
                            r'[Aa]g[:\s]+(\d{3,5})',          # Ag: 6677
                        ]
                        
                        melhor_conta = None
                        for pattern in conta_patterns:
                            match = re.search(pattern, credited_section)
                            if match:
                                melhor_conta = match.group(1).strip()
                                break
                        
                        melhor_agencia = None
                        for pattern in agencia_patterns:
                            match = re.search(pattern, credited_section)
                            if match:
                                melhor_agencia = match.group(1).strip()
                                break
                        
                        # Se não encontrou conta ou agência, pular
                        if not melhor_conta or not melhor_agencia:
                            continue
                        
                        # Normalizar conta e agência encontradas
                        conta_norm = normalize_account(melhor_conta)
                        agencia_norm = normalize_account(melhor_agencia)
                        
                        # Filtrar contas válidas (5-7 dígitos após normalização - contas geralmente têm 5+ dígitos)
                        if not conta_norm or len(conta_norm) < 5 or len(conta_norm) > 7:
                            continue
                        
                        # Filtrar agências válidas (3-5 dígitos)
                        if not agencia_norm or len(agencia_norm) < 3 or len(agencia_norm) > 5:
                            continue
                        
                        # Criar chave combinada conta+agência
                        chave_pdf = f"{conta_norm}_{agencia_norm}"
                        # Também criar chave invertida (caso na planilha esteja conta<->agência trocados)
                        chave_pdf_invertida = f"{agencia_norm}_{conta_norm}"
                        
                        # Verificar se a combinação conta+agência está na planilha
                        # Considera: combinação normal, combinação invertida, ou conta isolada
                        esta_cadastrado = (
                            chave_pdf in contas_excel_set or 
                            chave_pdf_invertida in contas_excel_invertido_set or
                            conta_norm in contas_excel_conta_set
                        )
                        
                        if not esta_cadastrado:
                            # Extrair um trecho do texto ao redor DA SEÇÃO CREDITADA
                            pos = credited_section.find(melhor_conta)
                            if pos != -1:
                                start = max(0, pos - 80)
                                end = min(len(credited_section), pos + 150)
                                snippet = credited_section[start:end].replace('\n', ' ')
                                snippet = ' '.join(snippet.split())
                                if len(snippet) > 200:
                                    snippet = snippet[:200] + "..."
                            else:
                                snippet = ' '.join(credited_section.split())[:200] + "..."
                            
                            nome_pdf = extract_name_from_page(page_data) or 'N/A'
                            nao_encontrados.append({
                                'pdf': pdf_name,
                                'pagina': page_num + 1,
                                'nome': nome_pdf,
                                'conta': melhor_conta,
                                'agencia': melhor_agencia,
                                'conta_normalizada': conta_norm,
                                'agencia_normalizada': agencia_norm,
                                'trecho': snippet
                            })
                
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao analisar {pdf_name}: {e}")
                    continue

            # Gerar arquivo TXT com comprovantes que NÃO têm funcionário na planilha
            if nao_encontrados:
                try:
                    txt_path = os.path.join(out_dir, f"comprovantes_sem_funcionario_{time.strftime('%Y%m%d_%H%M%S')}.txt")
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write("="*80 + "\n")
                        f.write("RELATÓRIO DE COMPROVANTES SEM FUNCIONÁRIO NA PLANILHA\n")
                        f.write("="*80 + "\n")
                        f.write(f"Data/Hora: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")
                        f.write(f"PDFs processados: {len(pdf_files)}\n")
                        f.write(f"Comprovantes extraídos com sucesso: {total_ok}\n")
                        f.write(f"Comprovantes SEM funcionário na planilha: {len(nao_encontrados)}\n")
                        f.write("="*80 + "\n\n")
                        f.write("ESTES SÃO COMPROVANTES QUE EXISTEM NOS PDFs MAS NÃO TÊM\n")
                        f.write("FUNCIONÁRIO CORRESPONDENTE CADASTRADO NA PLANILHA:\n")
                        f.write("-"*80 + "\n\n")

                        for idx, item in enumerate(nao_encontrados, 1):
                            f.write(f"{idx}. PDF: {item['pdf']}\n")
                            f.write(f"   Página: {item['pagina']}\n")
                            f.write(f"   Nome: {item.get('nome', 'N/A')}\n")
                            f.write(f"   Conta encontrada: {item['conta']}\n")
                            f.write(f"   Agência encontrada: {item.get('agencia', 'N/A')}\n")
                            f.write(f"   Status: Conta ou Agência NÃO cadastrada na planilha\n")
                            f.write("-"*80 + "\n\n")
                        
                        f.write("\n" + "="*80 + "\n")
                        f.write("O QUE FAZER:\n")
                        f.write("="*80 + "\n")
                        f.write("1. Verifique se estas contas deveriam estar cadastradas na planilha\n")
                        f.write("2. Adicione os funcionários faltantes na planilha se necessário\n")
                        f.write("3. Ou ignore se forem contas inválidas/irrelevantes\n")
                        f.write("4. Reprocesse após atualizar a planilha\n")
                        f.write("="*80 + "\n")

                    self.write_log(f"📄 Relatório salvo: {os.path.basename(txt_path)}")
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao gerar relatório: {e}")
            
            self.write_log("\n" + "="*50)
            self.write_log("📊 RESUMO DO PROCESSAMENTO")
            self.write_log("="*50)
            self.write_log(f"📂 PDFs processados: {len(novos_pdfs)}")
            self.write_log(f"📄 Total de páginas/comprovantes: {total_paginas_pdfs}")
            self.write_log(f"")
            self.write_log(f"✓ Comprovantes extraídos (com match): {total_ok} páginas")
            self.write_log(f"✗ Comprovantes SEM cadastro: {len(nao_encontrados)} páginas")
            self.write_log(f"❓ Outras páginas: {total_paginas_pdfs - total_ok - len(nao_encontrados)}")
            self.write_log(f"")
            if nao_encontrados:
                self.write_log(f"📝 Relatório de páginas sem funcionário salvo em TXT")
            if total_duplicates > 0:
                self.write_log(f"⚠️ Comprovantes em múltiplas páginas: {total_duplicates}")
            self.write_log(f"⏱️ Tempo total: {time_str}")
            self.write_log("="*50)
            
            # Mensagem de conclusão
            outras = total_paginas_pdfs - total_ok - len(nao_encontrados)

            # Garantir que a variável esteja inicializada antes de concatenar
            msg_resultado = ""
            msg_resultado += f"📄 Total de páginas: {total_paginas_pdfs}\n"
            msg_resultado += f"✓ Extraídos: {total_ok}\n"
            msg_resultado += f"✗ Sem funcionário: {len(nao_encontrados)}\n"
            if outras > 0:
                msg_resultado += f"❓ Outras: {outras}\n"
            if nao_encontrados:
                msg_resultado += f"📄 Ver relatório TXT\n"
            msg_resultado += f"⏱️ {time_str}"

            # Capturar as strings agora (evita capturar variáveis de escopo que podem não existir quando o lambda for executado)
            status_text = f"{total_ok}/{total_paginas_pdfs} extraídos"
            final_message = msg_resultado
            
            # Salvar estatísticas do último processamento para possível upload
            self.last_output_folder = out_dir
            self.last_process_stats = {
                'total_files': total_ok,
                'total_pages': total_paginas_pdfs,
                'success': total_ok > 0,
                'out_dir': out_dir
            }
            
            self.root.after(0, lambda s=status_text: self.status_var.set(s))
            self.root.after(0, lambda m=final_message: messagebox.showinfo("Concluído", m))

            
        except Exception as e:
            self.stop_timer()
            self.write_log(f"\n❌ ERRO: {e}")
            import traceback
            traceback.print_exc()
            # Capturar a mensagem de erro em variável local para o lambda
            err_msg = str(e)
            self.root.after(0, lambda m=err_msg: messagebox.showerror("Erro", m))
        finally:
            # Limpar cache de PDFs para liberar memória
            if hasattr(self, '_pdf_cache'):
                self._pdf_cache.clear()
            
            self.root.after(0, self.finish)
    
    def finish(self):
        self.prog.stop()
        self.prog.set(0)
        self.btn.configure(state='normal')
        self.status_var.set("Pronto")

        if self.last_process_stats and self.last_process_stats.get('success'):
            self.write_log(f"\n💡 Dica: Você pode enviar os comprovantes para o Google Drive")
            if not hasattr(self, 'upload_btn'):
                self.upload_btn = ctk.CTkButton(self.controls_frame,
                                                text="📤 Enviar para Drive",
                                                command=self.open_drive_upload_dialog,
                                                fg_color=self.ACCENT, hover_color=self.ACCENT_HOVER,
                                                width=160)
                self.upload_btn.pack(after=self.btn, side=tk.LEFT, padx=(0, 15))


if __name__ == "__main__":
    root = ctk.CTk()
    App(root)
    root.mainloop()