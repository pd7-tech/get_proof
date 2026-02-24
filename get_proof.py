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
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    from PIL import Image, ImageTk
except ImportError:
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
        # PIL not available, will work without logo
        Image = None
        ImageTk = None
    except ImportError:
        print("Erro: tkinter não instalado")
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
        
        # Criar janela
        self.window = tk.Toplevel(parent)
        self.window.title("📤 Enviar para Google Drive")
        self.window.transient(parent)
        self.window.grab_set()
        
        self.setup_ui()
        
        # Configurar tamanho e centralizar APÓS adicionar todo o conteúdo
        self.window.update_idletasks()
        self.window.geometry("1600x900")
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1600 // 2)
        y = (self.window.winfo_screenheight() // 2) - (900 // 2)
        self.window.geometry(f"1600x900+{x}+{y}")
    
    def setup_ui(self):
        # Container principal
        main = ttk.Frame(self.window, padding=20)
        main.pack(fill=tk.BOTH, expand=True)
        
        # CABEÇALHO
        header_frame = ttk.Frame(main)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, 
                 text="📤 Enviar Comprovantes para Google Drive",
                 font=('Segoe UI', 16, 'bold'),
                 foreground=self.app.colors['primary_blue']).pack()
        
        ttk.Label(header_frame,
                 text="Revise os arquivos e selecione o destino antes de enviar",
                 font=('Segoe UI', 10),
                 foreground=self.app.colors['dark_gray']).pack()
        
        # RESUMO
        summary_frame = ttk.LabelFrame(main, text="📊 Resumo", padding=15)
        summary_frame.pack(fill=tk.X, pady=(0, 15))
        
        info_text = f"""📁 Pasta origem: {os.path.basename(self.source_folder)}
📄 Total de arquivos: {self.file_summary['total_files']}
📂 Centros de custo: {self.file_summary['total_folders']}
💾 Tamanho total: {self.app.format_size(self.file_summary['total_size'])}"""
        
        ttk.Label(summary_frame, text=info_text, justify=tk.LEFT).pack(anchor=tk.W)
        
        # LISTA DE PASTAS
        list_frame = ttk.LabelFrame(main, text="📋 Arquivos por Centro de Custo", padding=15)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # TreeView
        tree_container = ttk.Frame(list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(tree_container, 
                           columns=('files', 'size'),
                           show='tree headings',
                           selectmode='none')
        
        tree.heading('#0', text='Centro de Custo')
        tree.heading('files', text='Arquivos')
        tree.heading('size', text='Tamanho')
        
        tree.column('#0', width=400)
        tree.column('files', width=100, anchor=tk.CENTER)
        tree.column('size', width=150, anchor=tk.CENTER)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Popular tree
        for ccusto, data in sorted(self.file_summary['folders'].items()):
            tree.insert('', 'end', 
                       text=f"✓ {ccusto}",
                       values=(data['count'], self.app.format_size(data['size'])))
        
        # DESTINO
        dest_frame = ttk.LabelFrame(main, text="🎯 Destino no Google Drive", padding=15)
        dest_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.drive_path = tk.StringVar()
        
        # Tentar detectar Google Drive
        detected = self.app.detect_google_drive_folder()
        if detected:
            self.drive_path.set(detected)
            ttk.Label(dest_frame, 
                     text=f"✓ Google Drive detectado automaticamente",
                     foreground=self.app.colors['success'],
                     font=('Segoe UI', 9)).pack(anchor=tk.W, pady=(0, 5))
        
        path_frame = ttk.Frame(dest_frame)
        path_frame.pack(fill=tk.X)
        
        ttk.Label(path_frame, text="Pasta:").pack(side=tk.LEFT, padx=(0, 10))
        
        entry = ttk.Entry(path_frame, 
                         textvariable=self.drive_path,
                         font=('Segoe UI', 10))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        ttk.Button(path_frame,
                  text="📁 Procurar...",
                  command=self.select_drive_folder).pack(side=tk.LEFT)
        
        # OPÇÕES
        options_frame = ttk.LabelFrame(main, text="⚙️ Opções", padding=15)
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.keep_local = tk.BooleanVar(value=True)
        self.create_backup = tk.BooleanVar(value=False)
        self.open_after = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(options_frame,
                       text="✓ Manter cópia local após upload",
                       variable=self.keep_local).pack(anchor=tk.W, pady=2)
        
        ttk.Checkbutton(options_frame,
                       text="✓ Criar backup antes de enviar (.zip)",
                       variable=self.create_backup).pack(anchor=tk.W, pady=2)
        
        ttk.Checkbutton(options_frame,
                       text="✓ Abrir pasta do Drive após conclusão",
                       variable=self.open_after).pack(anchor=tk.W, pady=2)
        
        # BOTÕES
        button_frame = ttk.Frame(main)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame,
                  text="❌ Cancelar",
                  command=self.window.destroy).pack(side=tk.LEFT)
        
        ttk.Button(button_frame,
                  text="📂 Abrir Pasta Local",
                  command=self.open_local_folder).pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(button_frame,
                  text="📤 Enviar para Drive",
                  style='Accent.TButton',
                  command=self.start_upload).pack(side=tk.RIGHT)
    
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
        
        # Criar janela
        self.window = tk.Toplevel(parent)
        self.window.title("📤 Enviando para Google Drive...")
        self.window.geometry("650x350")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Centralizar
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (650 // 2)
        y = (self.window.winfo_screenheight() // 2) - (350 // 2)
        self.window.geometry(f"+{x}+{y}")
        
        self.setup_ui()
    
    def setup_ui(self):
        main = ttk.Frame(self.window, padding=30)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main,
                 text="📤 Enviando arquivos para Google Drive",
                 font=('Segoe UI', 14, 'bold'),
                 foreground=self.app.colors['primary_blue']).pack(pady=(0, 20))
        
        # Status
        self.status_label = ttk.Label(main,
                                     text="Preparando upload...",
                                     font=('Segoe UI', 11))
        self.status_label.pack(pady=(0, 15))
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main,
                                       length=550,
                                       mode='determinate')
        self.progress.pack(pady=(0, 10))
        
        # Porcentagem
        self.percent_label = ttk.Label(main,
                                      text="0%",
                                      font=('Segoe UI', 10, 'bold'),
                                      foreground=self.app.colors['primary_blue'])
        self.percent_label.pack()
        
        # Arquivo atual
        self.current_file = ttk.Label(main,
                                     text="",
                                     font=('Segoe UI', 9),
                                     foreground=self.app.colors['dark_gray'])
        self.current_file.pack(pady=(15, 5))
        
        # Estatísticas
        self.stats_label = ttk.Label(main,
                                    text="0 / 0 arquivos • 0 MB / 0 MB",
                                    font=('Segoe UI', 9))
        self.stats_label.pack(pady=(0, 5))
        
        # Tempo estimado
        self.time_label = ttk.Label(main,
                                   text="Calculando tempo restante...",
                                   font=('Segoe UI', 9),
                                   foreground=self.app.colors['dark_gray'])
        self.time_label.pack()
        
        # Botões
        button_frame = ttk.Frame(main)
        button_frame.pack(pady=(25, 0))
        
        self.cancel_btn = ttk.Button(button_frame,
                                    text="❌ Cancelar",
                                    command=self.cancel)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def update_progress(self, current, total, current_file, bytes_sent, bytes_total, elapsed_time):
        """Atualiza o progresso do upload"""
        if self.cancelled:
            return False
        
        try:
            # Porcentagem
            percent = (current / total) * 100 if total > 0 else 0
            self.progress['value'] = percent
            self.percent_label.config(text=f"{percent:.1f}%")
            
            # Status
            self.status_label.config(text=f"Enviando arquivo {current} de {total}...")
            
            # Arquivo atual
            self.current_file.config(text=f"📄 {os.path.basename(current_file)}")
            
            # Estatísticas
            mb_sent = bytes_sent / (1024 * 1024)
            mb_total = bytes_total / (1024 * 1024)
            self.stats_label.config(
                text=f"{current} / {total} arquivos • {mb_sent:.1f} MB / {mb_total:.1f} MB")
            
            # Tempo restante
            if current > 0 and elapsed_time > 0:
                avg_time_per_file = elapsed_time / current
                remaining_files = total - current
                remaining_time = avg_time_per_file * remaining_files
                
                if remaining_time < 60:
                    time_str = f"~{int(remaining_time)}s restantes"
                elif remaining_time < 3600:
                    time_str = f"~{int(remaining_time / 60)}m restantes"
                else:
                    time_str = f"~{int(remaining_time / 3600)}h restantes"
                
                self.time_label.config(text=time_str)
            
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
            self.status_label.config(text="❌ Cancelando...")
            self.cancel_btn.config(state='disabled')
    
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
        
        # Criar janela
        self.window = tk.Toplevel(parent)
        self.window.title("✅ Upload Concluído" if results['success'] > 0 else "⚠️ Upload com Problemas")
        self.window.geometry("700x600")
        self.window.transient(parent)
        
        # Centralizar
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.window.winfo_screenheight() // 2) - (600 // 2)
        self.window.geometry(f"+{x}+{y}")
        
        self.setup_ui()
    
    def setup_ui(self):
        main = ttk.Frame(self.window, padding=30)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Ícone e título
        header = ttk.Frame(main)
        header.pack(pady=(0, 20))
        
        if self.results['errors'] == 0:
            icon_text = "✅"
            title_text = "Upload Concluído com Sucesso!"
            color = self.app.colors['success']
        else:
            icon_text = "⚠️"
            title_text = "Upload Concluído com Avisos"
            color = self.app.colors['warning']
        
        ttk.Label(header,
                 text=icon_text,
                 font=('Segoe UI', 48)).pack()
        
        ttk.Label(header,
                 text=title_text,
                 font=('Segoe UI', 16, 'bold'),
                 foreground=color).pack()
        
        # Estatísticas
        stats_frame = ttk.LabelFrame(main, text="📊 Estatísticas", padding=20)
        stats_frame.pack(fill=tk.X, pady=(0, 15))
        
        stats_text = f"""✓ {self.results['success']} arquivo(s) enviado(s) com sucesso
✗ {self.results['errors']} erro(s)
⏱️ Tempo total: {self.results['duration']}
💾 Dados transferidos: {self.results['size_mb']} MB
🔗 Destino: {os.path.basename(self.results['drive_url'])}"""
        
        ttk.Label(stats_frame, text=stats_text, justify=tk.LEFT).pack(anchor=tk.W)
        
        # Erros (se houver)
        if self.results['errors'] > 0 and self.results.get('error_list'):
            error_frame = ttk.LabelFrame(main, text="⚠️ Arquivos com Erro", padding=15)
            error_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
            
            # Lista de erros
            error_text = scrolledtext.ScrolledText(error_frame, 
                                                  height=10, 
                                                  font=('Consolas', 9))
            error_text.pack(fill=tk.BOTH, expand=True)
            
            for idx, error in enumerate(self.results['error_list'], 1):
                error_text.insert(tk.END, f"{idx}. {os.path.basename(error['file'])}\n")
                error_text.insert(tk.END, f"   Erro: {error['error']}\n\n")
            
            error_text.config(state='disabled')
        
        # Ações
        action_frame = ttk.Frame(main)
        action_frame.pack(pady=(10, 0))
        
        ttk.Button(action_frame,
                  text="🔗 Abrir no Drive",
                  command=lambda: self.open_drive(self.results['drive_url'])).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(action_frame,
                  text="📄 Salvar Relatório",
                  command=self.save_report).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(action_frame,
                  text="✓ Fechar",
                  style='Accent.TButton',
                  command=self.window.destroy).pack(side=tk.LEFT, padx=5)
    
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
    def __init__(self, root):
        self.root = root
        self.root.title("PD7Lab - Extrator de Comprovantes PDF v1.0.0")
        self.root.geometry("950x750")
        self.root.minsize(850, 650)
        
        # Tentar definir ícone da janela
        try:
            icon_path = resource_path("pd7-escudo.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                # Tentar com PNG se ICO não existir
                icon_png = resource_path("pd7-escudo.ico")
                if os.path.exists(icon_png):
                    icon_image = Image.open(icon_png)
                    icon_photo = ImageTk.PhotoImage(icon_image)
                    self.root.iconphoto(True, icon_photo)
        except Exception as e:
            # Continuar mesmo se não conseguir carregar o ícone
            pass
        
        # PD7Lab Color Palette (from logo)
        self.colors = {
            'primary_blue': '#00D4FF',      # Cyan blue from logo
            'dark_blue': '#0099CC',         # Darker shade
            'accent_blue': '#0AEAFF',       # Lighter cyan
            'white': '#F8F8F8',             # Off-white (slightly darker)
            'light_gray': '#F5F5F5',
            'medium_gray': '#E0E0E0',
            'dark_gray': '#424242',
            'text_dark': '#212121',
            'success': '#4CAF50',
            'warning': '#FF9800',
            'error': '#F44336'
        }
        
        # Set window background
        self.root.configure(bg=self.colors['white'])
        
        self.pdf_folder_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.out_var = tk.StringVar(value="comprovantes_extraidos")
        self.df = None
        self.conta_col = None
        self.agencia_col = None  # Nova coluna de agência
        self.nome_col = None
        self.ccusto_col = None
        self.last_dir = os.path.expanduser("~")
        
        # Option to force reprocess (ignore history)
        self.force_reprocess_var = tk.BooleanVar(value=False)
        
        # Debug mode - mostra detalhes de busca
        self.debug_mode_var = tk.BooleanVar(value=False)
        
        # Timer
        self.start_time = None
        self.timer_running = False
        self.timer_label = None
        
        # Logo image
        self.logo_image = None
        self.logo_label = None
        
        # Theme management
        self.current_theme = 'light'  # 'light' or 'dark'
        self.themes = {
            'light': {
                'primary_blue': '#00A8CC',      # Azul mais suave
                'dark_blue': '#008299',         # Tom mais profundo
                'accent_blue': '#00C4E6',       # Azul claro mais suave
                'white': '#F5F5F5',             # Cinza muito claro (não branco puro)
                'light_gray': '#E8E8E8',        # Cinza claro suave
                'medium_gray': '#CCCCCC',       # Cinza médio suave
                'dark_gray': '#5A5A5A',         # Cinza escuro mais suave
                'text_dark': '#303030',         # Texto cinza escuro (não preto puro)
                'success': '#43A047',           # Verde mais suave
                'warning': '#FB8C00',           # Laranja mais suave
                'error': '#E53935',             # Vermelho mais suave
                'logo_file': 'pd7lab-dark.jpeg'
            },
            'dark': {
                'primary_blue': '#00D4FF',
                'dark_blue': '#0099CC',
                'accent_blue': '#0AEAFF',
                'white': '#2B2B2B',          # Cinza escuro suave (não preto puro)
                'light_gray': '#363636',     # Cinza médio-escuro
                'medium_gray': '#4A4A4A',    # Cinza médio
                'dark_gray': '#9E9E9E',      # Cinza claro suave (não muito brilhante)
                'text_dark': '#D0D0D0',      # Texto cinza claro (não branco puro)
                'success': '#66BB6A',        # Verde mais suave
                'warning': '#FFA726',        # Laranja mais suave
                'error': '#EF5350',          # Vermelho mais suave
                'logo_file': 'pd7.png'
            }
        }
        
        # Histórico de PDFs processados
        self.processed_pdfs_file = "pdfs_processados.json"
        self.processed_pdfs = self.load_processed_pdfs()
        
        # Controle de último processamento (para upload)
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
        # Alternar tema
        self.current_theme = 'dark' if self.current_theme == 'light' else 'light'
        
        # Aplicar cores do novo tema
        theme_colors = self.themes[self.current_theme]
        self.colors = theme_colors.copy()
        
        # Recriar a UI com o novo tema
        # Limpar widgets existentes
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Reconfigurar background da janela
        self.root.configure(bg=self.colors['white'])
        
        # Recriar UI
        self.setup_ui()
        
        # Log da mudança
        theme_name = 'Escuro' if self.current_theme == 'dark' else 'Claro'
        self.write_log(f"🎨 Tema alterado para: {theme_name}")
    
    def setup_ui(self):
        # Apply PD7Lab themed style
        try:
            style = ttk.Style(self.root)
            # Use clam theme as base for better customization
            try:
                style.theme_use("clam")
            except:
                pass
            
            # Configure colors based on PD7Lab palette
            style.configure('TLabel', 
                          font=('Segoe UI', 10), 
                          background=self.colors['white'],
                          foreground=self.colors['text_dark'])
            
            style.configure('TButton', 
                          font=('Segoe UI', 10),
                          borderwidth=1,
                          relief='flat',
                          background=self.colors['medium_gray'],
                          foreground=self.colors['text_dark'])
            style.map('TButton', 
                     background=[('active', self.colors['primary_blue']),
                               ('pressed', self.colors['dark_blue'])],
                     foreground=[('active', self.colors['white'])])
            
            # Header style with PD7Lab blue
            style.configure('Header.TLabel', 
                          font=('Segoe UI', 18, 'bold'),
                          background=self.colors['white'],
                          foreground=self.colors['primary_blue'])
            
            # Accent button with PD7Lab colors
            style.configure('Accent.TButton', 
                          font=('Segoe UI', 11, 'bold'),
                          borderwidth=0,
                          relief='flat',
                          background=self.colors['primary_blue'],
                          foreground=self.colors['white'],
                          padding=(20, 10))
            style.map('Accent.TButton', 
                     background=[('active', self.colors['accent_blue']),
                               ('pressed', self.colors['dark_blue'])],
                     foreground=[('active', self.colors['white']),
                               ('pressed', self.colors['white'])])
            
            # Frame styles
            style.configure('TFrame', background=self.colors['white'])
            style.configure('TLabelframe', 
                          background=self.colors['white'],
                          foreground=self.colors['dark_gray'],
                          borderwidth=2,
                          relief='groove')
            style.configure('TLabelframe.Label', 
                          font=('Segoe UI', 10, 'bold'),
                          background=self.colors['white'],
                          foreground=self.colors['primary_blue'])
            
            # Entry style
            style.configure('TEntry',
                          fieldbackground=self.colors['white'],
                          foreground=self.colors['text_dark'],
                          borderwidth=1)
            
            # Checkbutton style
            style.configure('TCheckbutton',
                          background=self.colors['white'],
                          foreground=self.colors['text_dark'])
            
            # Progressbar with PD7Lab blue
            style.configure('TProgressbar',
                          troughcolor=self.colors['medium_gray'],
                          background=self.colors['primary_blue'],
                          borderwidth=0,
                          thickness=20)
            
        except Exception as e:
            # Fallback if styling fails
            print(f"Style warning: {e}")
            pass

        # Main container with white background
        main = ttk.Frame(self.root, padding=(15, 15))
        main.pack(fill=tk.BOTH, expand=True)

        # Header with logo
        header_frame = ttk.Frame(main)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Try to load and display logo
        try:
            if Image and ImageTk:
                logo_filename = self.themes[self.current_theme]['logo_file']
                logo_path = resource_path(logo_filename)  # Usar resource_path()
                if os.path.exists(logo_path):
                    logo_img = Image.open(logo_path)
                    # Resize logo to fit header (height ~60px)
                    aspect_ratio = logo_img.width / logo_img.height
                    new_height = 60
                    new_width = int(new_height * aspect_ratio)
                    logo_img = logo_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    self.logo_image = ImageTk.PhotoImage(logo_img)
                    
                    self.logo_label = ttk.Label(header_frame, image=self.logo_image, background=self.colors['white'])
                    self.logo_label.pack(side=tk.LEFT, padx=(0, 15))
                else:
                    # Se logo não encontrada, apenas registrar no log (não quebrar a aplicação)
                    print(f"Logo não encontrada: {logo_path}")
        except Exception as e:
            print(f"Logo loading warning: {e}")
            pass
        
        # Header text
        header_text_frame = ttk.Frame(header_frame)
        header_text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        header = ttk.Label(header_text_frame, text="Extrator de Comprovantes PDF v1.0.0", style='Header.TLabel')
        header.pack(anchor=tk.W)
        
        subtitle = ttk.Label(header_text_frame, 
                           text="Automatize a extração de comprovantes bancários", 
                           font=('Segoe UI', 9, 'italic'),
                           foreground=self.colors['dark_gray'])
        subtitle.pack(anchor=tk.W)
        
        # Theme toggle button on the right side of header
        theme_btn_frame = ttk.Frame(header_frame)
        theme_btn_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        theme_icon = "🌙" if self.current_theme == 'light' else "☀️"
        theme_text = "Modo Escuro" if self.current_theme == 'light' else "Modo Claro"
        
        self.theme_btn = ttk.Button(theme_btn_frame, 
                                    text=f"{theme_icon} {theme_text}", 
                                    command=self.toggle_theme,
                                    width=15)
        self.theme_btn.pack()
        
        # Separator line
        separator = tk.Frame(main, height=2, bg=self.colors['primary_blue'])
        separator.pack(fill=tk.X, pady=(0, 15))
        
        # Files section with custom styling
        files = ttk.LabelFrame(main, text="📁 Arquivos", padding=15)
        files.pack(fill=tk.X, pady=(0, 10))

        # Layout: label | entry | button
        files.columnconfigure(1, weight=1)

        ttk.Label(files, text="Pasta PDFs:").grid(row=0, column=0, sticky=tk.W, padx=(4, 12), pady=8)
        pdf_entry = ttk.Entry(files, textvariable=self.pdf_folder_var, font=('Segoe UI', 10))
        pdf_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10), pady=8, ipady=4)
        pdf_entry.bind('<Return>', lambda e: self.validate_pdf_folder())
        ttk.Button(files, text="Procurar...", width=14, command=self.get_pdf_folder).grid(row=0, column=2, padx=(0,4), pady=8)

        ttk.Label(files, text="Planilha Excel:").grid(row=1, column=0, sticky=tk.W, padx=(4, 12), pady=8)
        excel_entry = ttk.Entry(files, textvariable=self.excel_var, font=('Segoe UI', 10))
        excel_entry.grid(row=1, column=1, sticky='ew', padx=(0, 10), pady=8, ipady=4)
        excel_entry.bind('<Return>', lambda e: self.validate_excel())
        ttk.Button(files, text="Procurar...", width=14, command=self.get_excel).grid(row=1, column=2, padx=(0,4), pady=8)

        ttk.Label(files, text="Pasta de Saída:").grid(row=2, column=0, sticky=tk.W, padx=(4, 12), pady=8)
        out_entry = ttk.Entry(files, textvariable=self.out_var, font=('Segoe UI', 10))
        out_entry.grid(row=2, column=1, sticky='ew', padx=(0, 10), pady=8, ipady=4)
        out_entry.bind('<Return>', lambda e: self.validate_out())
        ttk.Button(files, text="Procurar...", width=14, command=self.get_out).grid(row=2, column=2, padx=(0,4), pady=8)

        # Status / timer row with colored background
        status_row = tk.Frame(main, bg=self.colors['light_gray'], relief='flat', bd=0)
        status_row.pack(fill=tk.X, pady=(10, 8))
        
        status_inner = tk.Frame(status_row, bg=self.colors['light_gray'])
        status_inner.pack(fill=tk.X, padx=10, pady=8)
        
        self.timer_label = tk.Label(status_inner, 
                                    text="⏱️ Tempo: 00:00:00.000",
                                    font=('Segoe UI', 10, 'bold'),
                                    bg=self.colors['light_gray'],
                                    fg=self.colors['primary_blue'])
        self.timer_label.pack(side=tk.LEFT)

        # Options frame for reprocess controls
        options_frame = ttk.LabelFrame(main, text="⚙️ Opções de Processamento", padding=12)
        options_frame.pack(fill=tk.X, pady=(8, 10))
        
        try:
            chk = ttk.Checkbutton(options_frame, text="Ignorar histórico (forçar reprocessamento)", 
                                 variable=self.force_reprocess_var)
            chk.pack(side=tk.LEFT, padx=(4, 12))
            
            chk_debug = ttk.Checkbutton(options_frame, text="🔧 Debug", 
                                       variable=self.debug_mode_var)
            chk_debug.pack(side=tk.LEFT, padx=(0, 12))
            
            ttk.Button(options_frame, text="🗑️ Limpar Histórico", 
                      command=self.clear_processed_history, width=18).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(options_frame, text="🔍 Buscar Não Encontrados", 
                      command=self.search_missing, width=24).pack(side=tk.LEFT, padx=(6, 4))
        except Exception:
            # ignore if style/ttk not available
            pass

        # Process button and progress with enhanced styling
        controls = ttk.Frame(main)
        controls.pack(fill=tk.X, pady=(12, 8))
        
        # Salvar referência ao frame de controles para adicionar botão de upload depois
        self.controls_frame = controls
        
        # Main action button with PD7Lab styling
        self.btn = ttk.Button(controls, text="▶ PROCESSAR COMPROVANTES", 
                             command=self.start, style='Accent.TButton')
        self.btn.pack(side=tk.LEFT, padx=(0, 15))

        self.prog = ttk.Progressbar(controls, mode='indeterminate', length=400)
        self.prog.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))

        # Status label to the right
        self.status_var = tk.StringVar(value="Pronto")
        status_label = ttk.Label(controls, 
                                textvariable=self.status_var, 
                                font=('Segoe UI', 9, 'italic'),
                                foreground=self.colors['dark_gray'])
        status_label.pack(side=tk.LEFT)

        # Log area with styled frame
        logf = ttk.LabelFrame(main, text="📋 Log de Processamento", padding=10)
        logf.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Create text widget with custom colors
        self.log = scrolledtext.ScrolledText(logf, 
                                            height=12, 
                                            state='disabled', 
                                            font=('Consolas', 9),
                                            bg=self.colors['white'],
                                            fg=self.colors['text_dark'],
                                            relief='flat',
                                            borderwidth=1,
                                            highlightthickness=1,
                                            highlightbackground=self.colors['medium_gray'],
                                            wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True)
        
        # Configure text tags for colored log messages
        self.log.tag_config('success', foreground=self.colors['success'], font=('Consolas', 9, 'bold'))
        self.log.tag_config('error', foreground=self.colors['error'], font=('Consolas', 9, 'bold'))
        self.log.tag_config('warning', foreground=self.colors['warning'], font=('Consolas', 9, 'bold'))
        self.log.tag_config('info', foreground=self.colors['primary_blue'], font=('Consolas', 9, 'bold'))
    
    def update_timer(self):
        """Atualiza o cronômetro a cada 100ms"""
        if self.timer_running and self.start_time:
            elapsed = time.time() - self.start_time
            hours, remainder = divmod(int(elapsed), 3600)
            minutes, seconds = divmod(remainder, 60)
            milliseconds = int((elapsed % 1) * 1000)
            time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"
            self.timer_label.config(text=f"⏱️ Tempo: {time_str}")
            self.root.after(100, self.update_timer)
    
    def start_timer(self):
        """Inicia o cronômetro"""
        self.start_time = time.time()
        self.timer_running = True
        self.timer_label.config(text="⏱️ Tempo: 00:00:00.000")
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
            # Primeira leitura para detectar colunas
            self.df = pd.read_excel(path)
            cols = list(self.df.columns)
            
            # Auto-detectar colunas (hardcoded)
            self.conta_col = find_column(self.df, ['conta', 'account', 'conta corrente'])
            self.agencia_col = find_column(self.df, ['agencia', 'agência', 'ag', 'agency'])
            self.nome_col = find_column(self.df, ['nome social', 'nome', 'funcionario'])
            self.ccusto_col = find_column(self.df, ['descrição ccusto', 'descricao ccusto', 'descrição de ccusto', 'descricao de ccusto', 'desc ccusto', 'ccusto', 'centro de custo', 'setor'])
            
            # Reler o Excel forçando conta e agência como TEXTO para preservar zeros à esquerda
            dtype_dict = {}
            if self.conta_col:
                dtype_dict[self.conta_col] = str
            if self.agencia_col:
                dtype_dict[self.agencia_col] = str
            
            if dtype_dict:
                self.df = pd.read_excel(path, dtype=dtype_dict)
                self.write_log(f"ℹ️ Colunas Conta/Agência lidas como TEXTO (preserva zeros à esquerda)")
            
            self.write_log(f"Colunas: {len(cols)} | Registros: {len(self.df)}")
            self.write_log(f"✓ Detectadas: Conta={self.conta_col}, Agência={self.agencia_col}, Nome={self.nome_col}, CCusto={self.ccusto_col}")
        except Exception as e:
            self.write_log(f"Erro: {e}")
    
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
            self.log.config(state='normal')
            self.log.insert(tk.END, msg + "\n")
            self.log.see(tk.END)
            self.log.config(state='disabled')
            self.root.update()
        except Exception:
            # Fallback se a janela não estiver disponível
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
        search_win = tk.Toplevel(self.root)
        search_win.title("🔍 Busca Assistida - Comprovantes Não Encontrados")
        search_win.geometry("1000x700")
        
        # Frame principal
        main_frame = ttk.Frame(search_win, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Label(main_frame, text="Busca Assistida de Comprovantes", 
                          font=('Segoe UI', 14, 'bold'))
        header.pack(pady=(0, 10))
        
        # Info
        info_text = f"Total de comprovantes não encontrados: {len(missing_items)}\n"
        info_text += "Selecione um item e clique em 'Buscar' para procurar nos PDFs com critérios flexíveis."
        info_label = ttk.Label(main_frame, text=info_text, font=('Segoe UI', 9))
        info_label.pack(pady=(0, 10))
        
        # Frame para lista e detalhes
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Lista de não encontrados (esquerda)
        list_frame = ttk.LabelFrame(content_frame, text="📋 Não Encontrados", padding=5)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Treeview para lista
        columns = ('conta', 'nome', 'ccusto')
        tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        tree.heading('conta', text='Conta')
        tree.heading('nome', text='Nome')
        tree.heading('ccusto', text='Centro de Custo')
        tree.column('conta', width=100)
        tree.column('nome', width=250)
        tree.column('ccusto', width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Adicionar itens
        for item in missing_items:
            tree.insert('', tk.END, values=(
                item.get('conta', ''),
                item.get('nome', ''),
                item.get('ccusto', '')
            ))
        
        # Frame de resultados (direita)
        results_frame = ttk.LabelFrame(content_frame, text="🔍 Resultados da Busca", padding=5)
        results_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Texto para resultados
        results_text = scrolledtext.ScrolledText(results_frame, height=20, width=50, 
                                                 font=('Courier New', 9), state='disabled')
        results_text.pack(fill=tk.BOTH, expand=True)
        
        # Frame de botões
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        status_var = tk.StringVar(value="Selecione um item e clique em Buscar")
        status_label = ttk.Label(button_frame, textvariable=status_var, font=('Segoe UI', 9, 'italic'))
        status_label.pack(side=tk.LEFT, padx=(0, 10))
        
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
            results_text.config(state='normal')
            results_text.delete(1.0, tk.END)
            results_text.insert(tk.END, f"Buscando por:\n")
            results_text.insert(tk.END, f"  Conta: {conta}\n")
            results_text.insert(tk.END, f"  Nome: {nome}\n")
            results_text.insert(tk.END, f"  C.Custo: {ccusto}\n")
            results_text.insert(tk.END, f"\n{'='*50}\n\n")
            results_text.config(state='disabled')

            def worker():
                try:
                    matches = self.flexible_search(conta, nome, ccusto)
                except Exception as e:
                    matches = []
                    err = e
                else:
                    err = None

                def finish_ui():
                    # Atualizar resultados na thread principal
                    current_results['matches'] = matches
                    results_text.config(state='normal')
                    results_text.delete(1.0, tk.END)
                    if err:
                        results_text.insert(tk.END, f"❌ Erro durante a busca: {err}\n")
                        status_var.set("Erro na busca")
                    elif matches:
                        results_text.insert(tk.END, f"✓ Encontrados {len(matches)} possíveis matches:\n\n")
                        for i, match in enumerate(matches, 1):
                            results_text.insert(tk.END, f"{i}. PDF: {match['pdf']}\n")
                            results_text.insert(tk.END, f"   Página: {match['page'] + 1}\n")
                            results_text.insert(tk.END, f"   Critério: {match.get('criteria','?')}\n")
                            results_text.insert(tk.END, f"   Trecho:\n")
                            results_text.insert(tk.END, f"   {match.get('snippet','')}\n")
                            results_text.insert(tk.END, f"\n{'-'*50}\n\n")
                        status_var.set(f"Encontrados {len(matches)} possíveis matches - Revise e confirme")
                    else:
                        results_text.insert(tk.END, "❌ Nenhum match encontrado mesmo com busca flexível.\n\n")
                        results_text.insert(tk.END, "Dicas:\n")
                        results_text.insert(tk.END, "• Verifique se o nome está correto\n")
                        results_text.insert(tk.END, "• Verifique se a conta está correta\n")
                        results_text.insert(tk.END, "• Verifique se o comprovante está no PDF\n")
                        status_var.set("Nenhum match encontrado")

                    results_text.config(state='disabled')

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
        
        ttk.Button(button_frame, text="🔍 Buscar", command=search_selected, width=15).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="✓ Extrair Selecionados", command=extract_selected, width=20).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="❌ Fechar", command=search_win.destroy, width=15).pack(side=tk.RIGHT)
    
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
            messagebox.showerror("Erro", "Colunas não encontradas no Excel!\nVerifique se existem as colunas: Conta, Agência, Nome e Descrição Ccusto")
            return
        
        self.btn.config(state='disabled')
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
                            
                            nao_encontrados.append({
                                'pdf': pdf_name,
                                'pagina': page_num + 1,
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
        self.btn.config(state='normal')
        self.status_var.set("Pronto")
        
        # Se houve processamento com sucesso, mostrar botão para upload ao Google Drive
        if self.last_process_stats and self.last_process_stats.get('success'):
            # Criar mensagem no log
            self.write_log(f"\n💡 Dica: Você pode enviar os comprovantes para o Google Drive")
            
            # Adicionar botão de upload (se ainda não existir)
            if not hasattr(self, 'upload_btn'):
                self.upload_btn = ttk.Button(self.controls_frame,
                                            text="📤 Enviar para Drive",
                                            command=self.open_drive_upload_dialog,
                                            width=20)
                # Inserir após o botão principal
                self.upload_btn.pack(after=self.btn, side=tk.LEFT, padx=(0, 15))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()