import os
import re
import sys
import threading
import time
from pathlib import Path
from datetime import timedelta
import shutil
import subprocess
import platform

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
except ImportError:
    print("Erro: tkinter não instalado")
    sys.exit(1)


def normalize_account(conta):
    """Normaliza conta removendo caracteres. Ex: '52938-2' -> '529382'"""
    if conta is None:
        return ""
    return re.sub(r'[^0-9]', '', str(conta))


def extract_pdf_pages(pdf_path):
    """Extrai texto de cada página do PDF"""
    pages = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages[i] = {'text': text, 'numbers': normalize_account(text)}
    return pages


def find_account_pages(conta, nome, pages):
    """Busca páginas onde TANTO a conta QUANTO o nome aparecem juntos"""
    found = []
    conta_norm = normalize_account(conta)
    conta_original = str(conta).strip()
    nome_upper = str(nome).upper().strip() if nome else ""
    
    if not conta_norm or len(conta_norm) < 3:
        return found
    
    if not nome_upper:
        return found
    
    # Para cada página, verifica se tem TANTO a conta QUANTO o nome
    for num, data in pages.items():
        text_upper = data['text'].upper()
        tem_conta = False
        tem_nome = False
        
        # Verifica se tem a conta (busca 1: com formatação)
        if conta_original in data['text']:
            tem_conta = True
        # Busca 2: conta normalizada
        elif conta_norm in data['numbers']:
            tem_conta = True
        # Busca 3: sem dígito verificador (último recurso)
        elif len(conta_norm) > 4:
            conta_sem_dv = conta_norm[:-1]
            if len(conta_sem_dv) >= 4 and conta_sem_dv in data['numbers']:
                tem_conta = True
        
        # Verifica se tem o nome (pode ser parcial para nomes compostos)
        if nome_upper in text_upper:
            tem_nome = True
        else:
            # Tenta verificar partes do nome (min 3 caracteres por parte)
            partes_nome = [p for p in nome_upper.split() if len(p) >= 3]
            if partes_nome and all(parte in text_upper for parte in partes_nome):
                tem_nome = True
        
        # SÓ adiciona se encontrou AMBOS: conta E nome
        if tem_conta and tem_nome:
            if num not in found:
                found.append(num)
    
    return found


def create_pdf(pdf_path, page_numbers, output_path):
    """Cria PDF com páginas específicas"""
    if not page_numbers:
        return False
    
    reader = None
    writer = None
    
    try:
        # Abrir o arquivo PDF fonte
        reader = PyPDF2.PdfReader(pdf_path)
        
        # Criar um novo writer para cada arquivo
        writer = PyPDF2.PdfWriter()
        
        # Adicionar apenas as páginas especificadas
        for num in page_numbers:
            if num < len(reader.pages):
                page = reader.pages[num]
                writer.add_page(page)
        
        # Verificar se há páginas e salvar
        if len(writer.pages) > 0:
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

            # escrever em arquivo temporário e mover para destino (atomicidade)
            try:
                import tempfile
                dirpath = os.path.dirname(target) or '.'
                fd, tmpname = tempfile.mkstemp(dir=dirpath, suffix='.pdf')
                os.close(fd)
                with open(tmpname, 'wb') as out:
                    writer.write(out)
                os.replace(tmpname, target)
            except Exception:
                # fallback simples
                with open(target, 'wb') as out:
                    writer.write(out)

            return True
            
    except Exception as e:
        print(f"Erro criar PDF: {e}")
        return False
    finally:
        # Limpar referências
        writer = None
        reader = None
    
    return False


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
        self.root.title("Extrator de Comprovantes PDF")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        self.pdf_folder_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.out_var = tk.StringVar(value="comprovantes_extraidos")
        self.df = None
        self.conta_col = None
        self.nome_col = None
        self.ccusto_col = None
        self.last_dir = os.path.expanduser("~")
        
        # Option to force reprocess (ignore history)
        self.force_reprocess_var = tk.BooleanVar(value=False)
        
        # Timer
        self.start_time = None
        self.timer_running = False
        self.timer_label = None
        
        # Histórico de PDFs processados
        self.processed_pdfs_file = "pdfs_processados.json"
        self.processed_pdfs = self.load_processed_pdfs()
        
        self.setup_ui()
    
    def load_processed_pdfs(self):
        """Carrega lista de PDFs já processados"""
        try:
            if os.path.exists(self.processed_pdfs_file):
                import json
                with open(self.processed_pdfs_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except:
            pass
        return {}
    
    def save_processed_pdfs(self):
        """Salva lista de PDFs processados"""
        try:
            import json
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
    
    def setup_ui(self):
        # Apply a clean ttk style
        try:
            style = ttk.Style(self.root)
            # Prefer a neutral theme if available
            for t in ("clam", "alt", "default"):
                try:
                    style.theme_use(t)
                    break
                except Exception:
                    pass
            style.configure('TLabel', font=('Segoe UI', 10))
            style.configure('TButton', font=('Segoe UI', 10))
            style.configure('Header.TLabel', font=('Segoe UI', 16, 'bold'))
            style.configure('Accent.TButton', font=('Segoe UI', 11, 'bold'), foreground='#ffffff', background='#0078D7')
            style.map('Accent.TButton', background=[('active', '#005A9E')])
        except Exception:
            # ignore style errors on restricted environments
            pass

        # Main container
        main = ttk.Frame(self.root, padding=(12, 12))
        main.pack(fill=tk.BOTH, expand=True)

        # Header
        header = ttk.Label(main, text="Extrator de Comprovantes", style='Header.TLabel')
        header.pack(pady=(6, 12))

        # Files group (grid layout for neat alignment)
        files = ttk.LabelFrame(main, text="📁 Arquivos", padding=12)
        files.pack(fill=tk.X, pady=6)

        # Layout: label | entry | button
        files.columnconfigure(1, weight=1)

        ttk.Label(files, text="Pasta PDFs:").grid(row=0, column=0, sticky=tk.W, padx=(4, 8), pady=6)
        pdf_entry = ttk.Entry(files, textvariable=self.pdf_folder_var)
        pdf_entry.grid(row=0, column=1, sticky='ew', padx=(0, 8), pady=6)
        pdf_entry.bind('<Return>', lambda e: self.validate_pdf_folder())
        ttk.Button(files, text="Procurar...", width=14, command=self.get_pdf_folder).grid(row=0, column=2, padx=(0,4), pady=6)

        ttk.Label(files, text="Planilha Excel:").grid(row=1, column=0, sticky=tk.W, padx=(4, 8), pady=6)
        excel_entry = ttk.Entry(files, textvariable=self.excel_var)
        excel_entry.grid(row=1, column=1, sticky='ew', padx=(0, 8), pady=6)
        excel_entry.bind('<Return>', lambda e: self.validate_excel())
        ttk.Button(files, text="Procurar...", width=14, command=self.get_excel).grid(row=1, column=2, padx=(0,4), pady=6)

        ttk.Label(files, text="Pasta de Saída:").grid(row=2, column=0, sticky=tk.W, padx=(4, 8), pady=6)
        out_entry = ttk.Entry(files, textvariable=self.out_var)
        out_entry.grid(row=2, column=1, sticky='ew', padx=(0, 8), pady=6)
        out_entry.bind('<Return>', lambda e: self.validate_out())
        ttk.Button(files, text="Procurar...", width=14, command=self.get_out).grid(row=2, column=2, padx=(0,4), pady=6)

        # Status / timer row
        status_row = ttk.Frame(main)
        status_row.pack(fill=tk.X, pady=(10,4))
        self.timer_label = ttk.Label(status_row, text="⏱️ Tempo: 00:00:00.000")
        self.timer_label.pack(side=tk.LEFT)

        # Options frame for reprocess controls
        options_frame = ttk.LabelFrame(main, text="⚙️ Opções de Processamento", padding=8)
        options_frame.pack(fill=tk.X, pady=(6,4))
        
        try:
            chk = ttk.Checkbutton(options_frame, text="Ignorar histórico (forçar reprocessamento de todos os PDFs)", 
                                 variable=self.force_reprocess_var)
            chk.pack(side=tk.LEFT, padx=(4,12))
            ttk.Button(options_frame, text="🗑️ Limpar Histórico", 
                      command=self.clear_processed_history, width=18).pack(side=tk.LEFT, padx=(0,4))
        except Exception:
            # ignore if style/ttk not available
            pass

        # Process button and progress
        controls = ttk.Frame(main)
        controls.pack(fill=tk.X, pady=(10,4))
        # Accent styled button (fall back to default if style not available)
        try:
            self.btn = ttk.Button(controls, text="▶ PROCESSAR COMPROVANTES", command=self.start, style='Accent.TButton')
        except Exception:
            self.btn = ttk.Button(controls, text="▶ PROCESSAR COMPROVANTES", command=self.start)
        self.btn.pack(side=tk.LEFT, padx=(0,10))

        self.prog = ttk.Progressbar(controls, mode='indeterminate', length=400)
        self.prog.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,10))

        # Status label to the right
        self.status_var = tk.StringVar(value="Pronto")
        status_label = ttk.Label(controls, textvariable=self.status_var, font=('Segoe UI', 9, 'italic'))
        status_label.pack(side=tk.LEFT)

        # Log area
        logf = ttk.LabelFrame(main, text="📋 Log de Processamento", padding=8)
        logf.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.log = scrolledtext.ScrolledText(logf, height=12, state='disabled', font=('Courier New', 10))
        self.log.pack(fill=tk.BOTH, expand=True)
    
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
                self.pdf_folder_var.set(folder)
                self.last_dir = folder
                try:
                    pdf_count = len([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
                except Exception:
                    pdf_count = 0
                self.write_log(f"✓ Pasta PDFs: {os.path.basename(folder)} ({pdf_count} PDFs)")
            else:
                self.write_log("ℹ️ Seleção de pasta cancelada pelo usuário.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar pasta: {e}")
    
    def get_excel(self):
        """Seleciona arquivo Excel usando explorador nativo do SO"""
        try:
            arquivo = self._native_select_file("Selecionar Planilha Excel", 
                                               [("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")])
            if arquivo:
                if os.path.isfile(arquivo):
                    self.excel_var.set(arquivo)
                    self.last_dir = os.path.dirname(arquivo)
                    self.write_log(f"✓ Excel: {os.path.basename(arquivo)}")
                    self.load_excel(arquivo)
                else:
                    self.write_log("⚠️ Arquivo selecionado não existe.")
                    messagebox.showwarning("Arquivo inválido", "O arquivo selecionado não existe.")
            else:
                self.write_log("ℹ️ Seleção de arquivo cancelada pelo usuário.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar Excel: {e}")
    
    def load_excel(self, path):
        try:
            self.df = pd.read_excel(path)
            cols = list(self.df.columns)
            
            # Auto-detectar colunas (hardcoded)
            self.conta_col = find_column(self.df, ['conta', 'account'])
            self.nome_col = find_column(self.df, ['nome social', 'nome', 'funcionario'])
            self.ccusto_col = find_column(self.df, ['descrição ccusto', 'descricao ccusto', 'descrição de ccusto', 'descricao de ccusto', 'desc ccusto', 'ccusto', 'centro de custo', 'setor'])
            
            self.write_log(f"Colunas: {len(cols)} | Registros: {len(self.df)}")
            self.write_log(f"✓ Detectadas: Conta={self.conta_col}, Nome={self.nome_col}, CCusto={self.ccusto_col}")
        except Exception as e:
            self.write_log(f"Erro: {e}")
    
    def get_out(self):
        """Seleciona pasta de saída usando explorador nativo do SO"""
        try:
            folder = self._native_select_folder("Selecionar Pasta de Saída")
            if folder:
                self.out_var.set(folder)
                self.last_dir = folder
                self.write_log(f"✓ Pasta de saída: {folder}")
            else:
                self.write_log("ℹ️ Seleção de pasta de saída cancelada pelo usuário.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar pasta: {e}")
    
    def _native_select_folder(self, title):
        """Seleciona pasta usando o explorador nativo do sistema operacional"""
        sistema = platform.system()
        
        # Linux - tentar zenity, kdialog, ou yad
        if sistema == "Linux":
            # Tentar zenity primeiro (GNOME)
            if shutil.which('zenity'):
                try:
                    result = subprocess.run(
                        ['zenity', '--file-selection', '--directory', f'--title={title}', f'--filename={self.last_dir}/'],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao usar zenity: {e}")
            
            # Tentar kdialog (KDE)
            if shutil.which('kdialog'):
                try:
                    result = subprocess.run(
                        ['kdialog', '--getexistingdirectory', self.last_dir, '--title', title],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao usar kdialog: {e}")
            
            # Tentar yad
            if shutil.which('yad'):
                try:
                    result = subprocess.run(
                        ['yad', '--file-selection', '--directory', f'--title={title}', f'--filename={self.last_dir}/'],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao usar yad: {e}")
        
        # Windows - usar powershell com FolderBrowserDialog
        elif sistema == "Windows":
            try:
                script = f'''
Add-Type -AssemblyName System.Windows.Forms
$folder = New-Object System.Windows.Forms.FolderBrowserDialog
$folder.Description = "{title}"
$folder.SelectedPath = "{self.last_dir.replace('/', '\\\\')}"
$folder.ShowNewFolderButton = $true
if ($folder.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    Write-Output $folder.SelectedPath
}}
'''
                result = subprocess.run(
                    ['powershell', '-NoProfile', '-Command', script],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout.strip()
            except Exception as e:
                self.write_log(f"⚠️ Erro ao usar explorador nativo Windows: {e}")
        
        # macOS - usar osascript
        elif sistema == "Darwin":
            try:
                script = f'choose folder with prompt "{title}" default location (POSIX file "{self.last_dir}")'
                result = subprocess.run(
                    ['osascript', '-e', script],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0 and result.stdout.strip():
                    # Converter formato macOS para POSIX
                    mac_path = result.stdout.strip()
                    if mac_path.startswith('alias '):
                        mac_path = mac_path[6:]
                    return mac_path.replace(':', '/')
            except Exception as e:
                self.write_log(f"⚠️ Erro ao usar explorador nativo macOS: {e}")
        
        # Fallback para tkinter (pode não ser nativo mas funciona em todos os SOs)
        self.write_log("ℹ️ Usando diálogo tkinter (explorador nativo não disponível)")
        return filedialog.askdirectory(initialdir=self.last_dir, title=title)
    
    def _native_select_file(self, title, filetypes):
        """Seleciona arquivo usando o explorador nativo do sistema operacional"""
        sistema = platform.system()
        
        # Linux - tentar zenity, kdialog, ou yad
        if sistema == "Linux":
            # Construir filtro para zenity
            filter_args = []
            for name, pattern in filetypes:
                if pattern != "*.*":
                    filter_args.extend(['--file-filter', f'{name} | {pattern}'])
            
            # Tentar zenity primeiro (GNOME)
            if shutil.which('zenity'):
                try:
                    cmd = ['zenity', '--file-selection', f'--title={title}', f'--filename={self.last_dir}/'] + filter_args
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao usar zenity: {e}")
            
            # Tentar kdialog (KDE)
            if shutil.which('kdialog'):
                try:
                    # Construir filtro para kdialog
                    filter_str = " ".join([pattern for _, pattern in filetypes if pattern != "*.*"])
                    result = subprocess.run(
                        ['kdialog', '--getopenfilename', self.last_dir, filter_str, '--title', title],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao usar kdialog: {e}")
            
            # Tentar yad
            if shutil.which('yad'):
                try:
                    cmd = ['yad', '--file-selection', f'--title={title}', f'--filename={self.last_dir}/'] + filter_args
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao usar yad: {e}")
        
        # Windows - usar powershell com OpenFileDialog
        elif sistema == "Windows":
            try:
                # Construir filtro de tipos
                filter_parts = []
                for name, pattern in filetypes:
                    if pattern != "*.*":
                        filter_parts.append(f"{name}|{pattern}")
                filter_str = "|".join(filter_parts) if filter_parts else "Todos os arquivos|*.*"
                
                script = f'''
Add-Type -AssemblyName System.Windows.Forms
$file = New-Object System.Windows.Forms.OpenFileDialog
$file.Title = "{title}"
$file.InitialDirectory = "{self.last_dir.replace('/', '\\\\')}"
$file.Filter = "{filter_str}"
if ($file.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    Write-Output $file.FileName
}}
'''
                result = subprocess.run(
                    ['powershell', '-NoProfile', '-Command', script],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout.strip()
            except Exception as e:
                self.write_log(f"⚠️ Erro ao usar explorador nativo Windows: {e}")
        
        # macOS - usar osascript
        elif sistema == "Darwin":
            try:
                # Construir filtro de tipos para macOS
                extensions = []
                for _, pattern in filetypes:
                    if pattern != "*.*":
                        exts = pattern.replace("*.", "").split()
                        extensions.extend([f'"{ext}"' for ext in exts])
                
                type_filter = f" of type {{{','.join(extensions)}}}" if extensions else ""
                script = f'choose file with prompt "{title}"{type_filter} default location (POSIX file "{self.last_dir}")'
                
                result = subprocess.run(
                    ['osascript', '-e', script],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0 and result.stdout.strip():
                    # Converter formato macOS para POSIX
                    mac_path = result.stdout.strip()
                    if mac_path.startswith('alias '):
                        mac_path = mac_path[6:]
                    return mac_path.replace(':', '/')
            except Exception as e:
                self.write_log(f"⚠️ Erro ao usar explorador nativo macOS: {e}")
        
        # Fallback para tkinter
        self.write_log("ℹ️ Usando diálogo tkinter (explorador nativo não disponível)")
        return filedialog.askopenfilename(initialdir=self.last_dir, title=title, filetypes=filetypes)
    
    def validate_pdf_folder(self):
        """Valida caminho da pasta de PDFs digitada"""
        path = self.pdf_folder_var.get().strip()
        if path and os.path.exists(path) and os.path.isdir(path):
            self.last_dir = path
            pdf_count = len([f for f in os.listdir(path) if f.lower().endswith('.pdf')])
            self.write_log(f"✓ Pasta PDFs: {os.path.basename(path)} ({pdf_count} PDFs)")
        elif path:
            messagebox.showwarning("Aviso", "Pasta não encontrada!")
    
    def validate_excel(self):
        """Valida caminho do Excel digitado"""
        path = self.excel_var.get().strip()
        if path and os.path.exists(path) and (path.endswith('.xlsx') or path.endswith('.xls')):
            self.last_dir = os.path.dirname(path)
            self.write_log(f"✓ Excel: {os.path.basename(path)}")
            self.load_excel(path)
        elif path:
            messagebox.showwarning("Aviso", "Arquivo Excel não encontrado!")
    
    def validate_out(self):
        """Valida pasta de saída"""
        path = self.out_var.get().strip()
        if path:
            self.write_log(f"✓ Pasta: {path}")
    
    def write_log(self, msg):
        self.log.config(state='normal')
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state='disabled')
        self.root.update()

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
    
    def start(self):
        if not self.pdf_folder_var.get() or not self.excel_var.get():
            messagebox.showerror("Erro", "Selecione a pasta de PDFs e o Excel!")
            return
        if self.df is None:
            messagebox.showerror("Erro", "Carregue Excel!")
            return
        if not self.conta_col or not self.nome_col or not self.ccusto_col:
            messagebox.showerror("Erro", "Colunas não encontradas no Excel!\nVerifique se existem as colunas: Conta, Nome e Descrição Ccusto")
            return
        
        self.btn.config(state='disabled')
        self.status_var.set("Processando...")
        self.prog.start()
        self.start_timer()
        threading.Thread(target=self.process, daemon=True).start()
    
    def process(self):
        try:
            pdf_folder = self.pdf_folder_var.get()
            out_dir = self.out_var.get()
            conta_col = self.conta_col
            nome_col = self.nome_col
            ccusto_col = self.ccusto_col
            
            Path(out_dir).mkdir(parents=True, exist_ok=True)
            
            self.write_log("\n" + "="*50)
            self.write_log("🚀 Iniciando processamento...")
            self.write_log("="*50)
            
            # Listar todos os PDFs na pasta
            pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
            pdf_files.sort()
            
            if not pdf_files:
                self.write_log("\n⚠️ Nenhum PDF encontrado na pasta!")
                return
            
            self.write_log(f"\n� Total de PDFs na pasta: {len(pdf_files)}")
            
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
                nome = row[nome_col]
                ccusto = row[ccusto_col]
                
                if pd.isna(conta) or str(conta).strip() == '':
                    continue
                
                conta_str = str(conta).strip()
                nome_str = str(nome).strip() if not pd.isna(nome) else 'N/A'
                ccusto_str = str(ccusto).strip() if not pd.isna(ccusto) else 'N/A'
                
                todas_contas.append({
                    'conta': conta_str,
                    'nome': nome_str,
                    'ccusto': ccusto_str
                })
            
            for idx, (pdf_name, pdf_path, fingerprint) in enumerate(novos_pdfs, 1):
                self.write_log(f"\n{'='*50}")
                self.write_log(f"📄 Processando PDF {idx}/{len(novos_pdfs)}: {pdf_name}")
                self.write_log(f"{'='*50}")
                self.root.after(0, lambda i=idx, t=len(novos_pdfs): self.status_var.set(f"PDF {i}/{t}..."))
                
                try:
                    pages = extract_pdf_pages(pdf_path)
                    self.write_log(f"✓ Páginas extraídas: {len(pages)}")
                    
                    ok = 0
                    nok = 0
                    duplicates = 0
                    
                    for row_idx, row in self.df.iterrows():
                        conta = row[conta_col]
                        nome = row[nome_col]
                        ccusto = row[ccusto_col]
                        
                        # Verificar se dados estão presentes
                        if pd.isna(conta) or str(conta).strip() == '':
                            continue
                        
                        conta_str = str(conta).strip()
                        nome_str = clean_filename(nome)
                        ccusto_str = clean_filename(ccusto)
                        
                        paginas = find_account_pages(conta_str, nome, pages)
                        
                        if paginas:
                            if len(paginas) > 1:
                                duplicates += 1
                                self.write_log(f"⚠️ Conta {conta_str} em {len(paginas)} páginas: {[p+1 for p in paginas]}")
                            
                            out = os.path.join(out_dir, f"{ccusto_str}_{nome_str}.pdf")
                            i = 1
                            while os.path.exists(out):
                                out = os.path.join(out_dir, f"{ccusto_str}_{nome_str}_{i}.pdf")
                                i += 1
                            
                            if create_pdf(pdf_path, paginas, out):
                                self.write_log(f"✓ {ccusto_str}_{nome_str} (pág {[p+1 for p in paginas]})")
                                ok += 1
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
                            'nao_encontrados': nok
                        }
                        self.save_processed_pdfs()
                    
                    total_ok += ok
                    total_nok += nok
                    total_duplicates += duplicates
                    
                    self.write_log(f"✓ PDF concluído: {ok} extraídos, {nok} não encontrados")
                    
                except Exception as e:
                    self.write_log(f"❌ Erro ao processar {pdf_name}: {e}")
            
            # Parar timer e calcular tempo total
            elapsed = self.stop_timer()
            time_str = self.format_time(elapsed)
            
            # Identificar contas que NÃO foram encontradas em NENHUM PDF
            nao_encontrados = []
            for conta_info in todas_contas:
                if conta_info['conta'] not in contas_encontradas:
                    nao_encontrados.append(conta_info)
            
            # Gerar arquivo TXT com comprovantes não encontrados
            if nao_encontrados:
                try:
                    txt_path = os.path.join(out_dir, f"nao_encontrados_{time.strftime('%Y%m%d_%H%M%S')}.txt")
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write("="*80 + "\n")
                        f.write("RELATÓRIO DE COMPROVANTES NÃO ENCONTRADOS\n")
                        f.write("="*80 + "\n")
                        f.write(f"Data/Hora: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")
                        f.write(f"Total de comprovantes não encontrados: {len(nao_encontrados)}\n")
                        f.write(f"Total de contas no Excel: {len(todas_contas)}\n")
                        f.write(f"Total de contas encontradas: {len(contas_encontradas)}\n")
                        f.write("="*80 + "\n\n")
                        
                        for idx, item in enumerate(nao_encontrados, 1):
                            f.write(f"{idx}. Conta: {item['conta']}\n")
                            f.write(f"   Nome: {item['nome']}\n")
                            f.write(f"   Centro de Custo: {item['ccusto']}\n")
                            f.write("-"*80 + "\n")
                        
                        f.write("\n" + "="*80 + "\n")
                        f.write("OBSERVAÇÃO: Estas contas NÃO foram encontradas em NENHUM dos PDFs processados.\n")
                        f.write("Verifique se os dados estão corretos ou se os PDFs contêm estas informações.\n")
                        f.write("="*80 + "\n")
                    
                    self.write_log(f"📄 Relatório de não encontrados salvo em: {os.path.basename(txt_path)}")
                except Exception as e:
                    self.write_log(f"⚠️ Erro ao gerar relatório TXT: {e}")
            
            self.write_log("\n" + "="*50)
            self.write_log("📊 RESUMO GERAL DO PROCESSAMENTO")
            self.write_log("="*50)
            self.write_log(f"📂 PDFs na pasta: {len(pdf_files)}")
            self.write_log(f"⏭️ Já processados: {len(ja_processados)}")
            self.write_log(f"🆕 Novos processados: {len(novos_pdfs)}")
            self.write_log(f"📊 Total de contas no Excel: {len(todas_contas)}")
            self.write_log(f"✓ Total extraídos: {total_ok}")
            self.write_log(f"✗ Total não encontrados: {len(nao_encontrados)}")
            if nao_encontrados:
                self.write_log(f"📝 Comprovantes sem match salvos em arquivo TXT: {len(nao_encontrados)}")
            if total_duplicates > 0:
                self.write_log(f"⚠️ Contas duplicadas: {total_duplicates}")
            self.write_log(f"⏱️ Tempo total: {time_str}")
            self.write_log("="*50)
            
            # Mensagem de conclusão
            msg_resultado = f"PDFs processados: {len(novos_pdfs)}/{len(pdf_files)}\n"
            msg_resultado += f"📊 Contas no Excel: {len(todas_contas)}\n"
            msg_resultado += f"✓ Extraídos: {total_ok}\n"
            msg_resultado += f"✗ Não encontrados: {len(nao_encontrados)}\n"
            if nao_encontrados:
                msg_resultado += f"\n📄 Relatório de não encontrados gerado!\n"
            msg_resultado += f"⏱️ Tempo: {time_str}"
            
            self.root.after(0, lambda: self.status_var.set(f"Concluído - {total_ok} extraídos, {len(nao_encontrados)} não encontrados"))
            self.root.after(0, lambda: messagebox.showinfo("Processamento Concluído", msg_resultado))
            
        except Exception as e:
            self.stop_timer()
            self.write_log(f"\n❌ ERRO: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            self.root.after(0, self.finish)
    
    def finish(self):
        self.prog.stop()
        self.btn.config(state='normal')
        self.status_var.set("Pronto")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
