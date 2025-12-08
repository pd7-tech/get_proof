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

            pages[i] = {
                'text': text,
                'numbers': normalize_account(text),
                'norm_text': normalize_search_text(text)
            }
    return pages


def find_account_pages(conta, nome, pages):
    """Busca páginas onde TANTO a conta QUANTO o nome aparecem juntos"""
    found = []
    conta_norm = normalize_account(conta)
    conta_original = str(conta).strip()
    # Normalizar nome para busca (remover acentos e pontuação)
    def normalize_search_text(s):
        if not s:
            return ""
        nf = unicodedata.normalize('NFKD', str(s))
        ascii_s = nf.encode('ascii', 'ignore').decode('ascii')
        cleaned = re.sub(r'[^A-Za-z0-9\s]', ' ', ascii_s)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip().upper()
        return cleaned

    nome_norm = normalize_search_text(nome) if nome else ""
    
    if not conta_norm or len(conta_norm) < 3:
        return found
    
    if not nome_norm:
        return found
    
    # Para cada página, verifica se tem TANTO a conta QUANTO o nome
    for num, data in pages.items():
        text_upper = data['text'].upper()
        text_norm = data.get('norm_text', '')
        tem_conta = False
        tem_nome = False
        
        # Verifica se tem a conta (busca 1: conta numérica normalizada no conjunto de dígitos)
        if conta_norm and conta_norm in data['numbers']:
            tem_conta = True
        else:
            # Tentar encontrar os dígitos da conta no texto permitindo separadores (ex: 12345-6 ou 12 345 6)
            if conta_norm:
                pattern = "\\D*".join(list(conta_norm))
                try:
                    if re.search(pattern, data['text'] or '', flags=re.DOTALL):
                        tem_conta = True
                except re.error:
                    pass
        # Busca alternativa: sem dígito verificador (último recurso)
        if not tem_conta and len(conta_norm) > 4:
            conta_sem_dv = conta_norm[:-1]
            if len(conta_sem_dv) >= 4 and conta_sem_dv in data['numbers']:
                tem_conta = True
        
        # Verifica se tem o nome (usar texto normalizado sem acentos)
        if nome_norm and nome_norm in text_norm:
            tem_nome = True
        else:
            # Tenta verificar partes do nome (min 3 caracteres por parte)
            partes_nome = [p for p in nome_norm.split() if len(p) >= 3]
            if partes_nome:
                # exigir pelo menos 2 partes quando o nome possui múltiplas partes, senão 1
                need = 2 if len(partes_nome) >= 2 else 1
                matches = sum(1 for parte in partes_nome if parte in text_norm)
                if matches >= need:
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

            # Salvar diretamente no arquivo de destino
            try:
                with open(target, 'wb') as out:
                    writer.write(out)
            except Exception as e:
                print(f"Erro ao salvar PDF {target}: {e}")
                return False

            return True
            
    except Exception as e:
        print(f"Erro criar PDF: {e}")
        return False
    finally:
        # Limpar referências
        writer = None
        reader = None
    
    return False


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
                      command=self.clear_processed_history, width=22).pack(side=tk.LEFT, padx=(0,4))
            ttk.Button(options_frame, text="🔍 Buscar Não Encontrados", 
                      command=self.search_missing, width=28).pack(side=tk.LEFT, padx=(4,4))
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
        """Seleciona pasta usando o explorador nativo do sistema operacional"""
        # Usar sempre o diálogo tkinter padrão (funciona bem no Windows)
        folder = filedialog.askdirectory(initialdir=self.last_dir, title=title)
        if folder:
            return normalize_path(folder)
        return None
    
    def _native_select_file(self, title, filetypes):
        """Seleciona arquivo usando o explorador nativo do sistema operacional"""
        # Usar sempre o diálogo tkinter padrão (funciona bem no Windows)
        arquivo = filedialog.askopenfilename(initialdir=self.last_dir, title=title, filetypes=filetypes)
        if arquivo:
            return normalize_path(arquivo)
        return None
    
    def validate_pdf_folder(self):
        """Valida caminho da pasta de PDFs digitada"""
        path = normalize_path(self.pdf_folder_var.get().strip())
        if path and os.path.exists(path) and os.path.isdir(path):
            self.last_dir = path
            try:
                # Tentar múltiplas abordagens para listar PDFs (útil para OneDrive)
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
        """Valida caminho do Excel digitado"""
        path = normalize_path(self.excel_var.get().strip())
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
            
            current_item = {}
            for line in lines:
                line = line.strip()
                if line.startswith("Conta:"):
                    if current_item:
                        items.append(current_item)
                    current_item = {'conta': line.split("Conta:", 1)[1].strip()}
                elif line.startswith("Nome:"):
                    current_item['nome'] = line.split("Nome:", 1)[1].strip()
                elif line.startswith("Centro de Custo:"):
                    current_item['ccusto'] = line.split("Centro de Custo:", 1)[1].strip()
            
            if current_item:
                items.append(current_item)
                
        except Exception as e:
            self.write_log(f"❌ Erro ao ler arquivo: {e}")
        
        return items
    
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
            """Busca o item selecionado nos PDFs"""
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
            
            status_var.set(f"Buscando: {nome}...")
            results_text.config(state='normal')
            results_text.delete(1.0, tk.END)
            results_text.insert(tk.END, f"Buscando por:\n")
            results_text.insert(tk.END, f"  Conta: {conta}\n")
            results_text.insert(tk.END, f"  Nome: {nome}\n")
            results_text.insert(tk.END, f"  C.Custo: {ccusto}\n")
            results_text.insert(tk.END, f"\n{'='*50}\n\n")
            results_text.config(state='disabled')
            search_win.update()
            
            # Buscar nos PDFs com critérios flexíveis
            matches = self.flexible_search(conta, nome, ccusto)
            current_results['matches'] = matches
            
            # Mostrar resultados
            results_text.config(state='normal')
            if matches:
                results_text.insert(tk.END, f"✓ Encontrados {len(matches)} possíveis matches:\n\n")
                for i, match in enumerate(matches, 1):
                    results_text.insert(tk.END, f"{i}. PDF: {match['pdf']}\n")
                    results_text.insert(tk.END, f"   Página: {match['page'] + 1}\n")
                    results_text.insert(tk.END, f"   Critério: {match['criteria']}\n")
                    results_text.insert(tk.END, f"   Trecho:\n")
                    results_text.insert(tk.END, f"   {match['snippet']}\n")
                    results_text.insert(tk.END, f"\n{'-'*50}\n\n")
                status_var.set(f"Encontrados {len(matches)} possíveis matches - Revise e confirme")
            else:
                results_text.insert(tk.END, "❌ Nenhum match encontrado mesmo com busca flexível.\n")
                results_text.insert(tk.END, "\nDicas:\n")
                results_text.insert(tk.END, "• Verifique se o nome está correto\n")
                results_text.insert(tk.END, "• Verifique se a conta está correta\n")
                results_text.insert(tk.END, "• Verifique se o comprovante está no PDF\n")
                status_var.set("Nenhum match encontrado")
            results_text.config(state='disabled')
        
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
                
                out_path = os.path.join(out_dir, f"{ccusto_str}_{nome_str}_manual.pdf")
                i = 1
                while os.path.exists(out_path):
                    out_path = os.path.join(out_dir, f"{ccusto_str}_{nome_str}_manual_{i}.pdf")
                    i += 1
                
                if create_pdf(pdf_path, [match['page']], out_path):
                    success_count += 1
                    self.write_log(f"✓ Extraído manualmente: {ccusto_str}_{nome_str} (pág {match['page'] + 1})")
            
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
            pdf_folder = normalize_path(self.pdf_folder_var.get())
            out_dir = normalize_path(self.out_var.get())
            conta_col = self.conta_col
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
            
            # Rastrear páginas processadas
            total_paginas_pdfs = 0
            paginas_com_match = set()  # páginas que tiveram match (PDF + número da página)
            
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
                        nome = row[nome_col]
                        ccusto = row[ccusto_col]
                        
                        # Verificar se dados estão presentes - TODOS os campos obrigatórios
                        if pd.isna(conta) or str(conta).strip() == '':
                            continue
                        if pd.isna(nome) or str(nome).strip() == '':
                            continue
                        if pd.isna(ccusto) or str(ccusto).strip() == '':
                            continue
                        
                        # Garantir que as variáveis são sempre recriadas para cada linha
                        conta_str = str(conta).strip()
                        nome_str = clean_filename(str(nome).strip())
                        ccusto_str = clean_filename(str(ccusto).strip())
                        
                        paginas = find_account_pages(conta_str, nome, pages)
                        
                        if paginas:
                            # Registrar quais páginas tiveram match
                            for pag in paginas:
                                paginas_com_match.add(f"{pdf_name}|{pag}")
                            
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
                    
                    self.write_log(f"✓ Comprovantes extraídos deste PDF: {ok}")
                    
                except Exception as e:
                    self.write_log(f"❌ Erro ao processar {pdf_name}: {e}")
            
            # Calcular quantas páginas dos PDFs ficaram SEM match com a planilha
            paginas_sem_match = total_paginas_pdfs - len(paginas_com_match)
            
            self.write_log(f"\n📊 ESTATÍSTICAS DE PÁGINAS:")
            self.write_log(f"   Total de páginas nos PDFs: {total_paginas_pdfs}")
            self.write_log(f"   Páginas COM match (extraídas): {len(paginas_com_match)}")
            self.write_log(f"   Páginas SEM match na planilha: {paginas_sem_match}")
            
            # Parar timer e calcular tempo total
            elapsed = self.stop_timer()
            time_str = self.format_time(elapsed)
            
            # Comprovantes nos PDFs que NÃO têm funcionário correspondente na planilha
            nao_encontrados = []
            
            # Criar índice de contas do Excel para busca rápida
            contas_excel_set = set()
            for conta_info in todas_contas:
                conta_norm = normalize_account(conta_info['conta'])
                if conta_norm:
                    contas_excel_set.add(conta_norm)
            
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
                        
                        text = page_data.get('text', '')
                        
                        # Procurar padrões de conta (4-6 dígitos, possivelmente com separador)
                        acc_pattern = re.compile(r'\b\d{4,6}[-\s]?\d?\b')
                        contas_encontradas_pagina = acc_pattern.findall(text)
                        
                        if not contas_encontradas_pagina:
                            # Página sem padrão de conta - pode ser página em branco ou cabeçalho
                            continue
                        
                        # Pegar a primeira conta mais provável (geralmente a principal da página)
                        melhor_conta = None
                        for conta_raw in contas_encontradas_pagina:
                            conta_norm = normalize_account(conta_raw)
                            
                            # Filtrar contas válidas (4-7 dígitos)
                            if conta_norm and len(conta_norm) >= 4 and len(conta_norm) <= 7:
                                melhor_conta = conta_raw
                                break
                        
                        if not melhor_conta:
                            continue
                        
                        conta_norm = normalize_account(melhor_conta)
                        
                        # Verificar se a conta NÃO está na planilha
                        if conta_norm not in contas_excel_set:
                            # Extrair um trecho do texto ao redor
                            pos = text.find(melhor_conta)
                            if pos != -1:
                                start = max(0, pos - 80)
                                end = min(len(text), pos + 150)
                                snippet = text[start:end].replace('\n', ' ')
                                snippet = ' '.join(snippet.split())
                                if len(snippet) > 200:
                                    snippet = snippet[:200] + "..."
                            else:
                                snippet = ' '.join(text.split())[:200] + "..."
                            
                            nao_encontrados.append({
                                'pdf': pdf_name,
                                'pagina': page_num + 1,
                                'conta': melhor_conta,
                                'conta_normalizada': conta_norm,
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
                            f.write(f"   Status: Conta NÃO cadastrada na planilha\n")
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
            self.write_log(f"✗ Comprovantes SEM funcionário cadastrado: {len(nao_encontrados)} páginas")
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
            
            msg_resultado += f"📄 Total de páginas: {total_paginas_pdfs}"
            msg_resultado += f"✓ Extraídos: {total_ok}"
            msg_resultado += f"✗ Sem funcionário: {len(nao_encontrados)}"
            if outras > 0:
                msg_resultado += f"❓ Outras: {outras}"
            if nao_encontrados:
                msg_resultado += f"📄 Ver relatório TXT"
            msg_resultado += f"⏱️ {time_str}"
            
            self.root.after(0, lambda: self.status_var.set(f"{total_ok}/{total_paginas_pdfs} extraídos"))
            self.root.after(0, lambda: messagebox.showinfo("Concluído", msg_resultado))

            
        except Exception as e:
            self.stop_timer()
            self.write_log(f"\n❌ ERRO: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            # Limpar cache de PDFs para liberar memória
            if hasattr(self, '_pdf_cache'):
                self._pdf_cache.clear()
            
            self.root.after(0, self.finish)
    
    def finish(self):
        self.prog.stop()
        self.btn.config(state='normal')
        self.status_var.set("Pronto")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()