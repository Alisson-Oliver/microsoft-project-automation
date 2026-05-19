import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import re
import ast

CONFIG_FILE = Path(__file__).parent / "config.py"

class ConfiguradorCompleto:
    def __init__(self, root):
        self.root = root
        self.root.title("Configurador do Sistema - MPP → XLSX")
        self.root.geometry("900x700")
        
        self.config_dados = self._carregar_config()
        self.colunas = self.config_dados.get('colunas', [])
        self.selecionada = None
        
        self._criar_widgets()
        self._atualizar_campos()
    
    def _carregar_config(self):
        """Carrega todas as configurações do config.py"""
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                conteudo = f.read()

            dados = {}

            def _get_val(var_name):
                m = re.search(rf'^{var_name}\s*=\s*(.+)$', conteudo, re.MULTILINE)
                if not m:
                    return None
                expr = m.group(1).strip()
                if expr.startswith(('r"', "r'", 'R"', "R'")):
                    expr = expr[1:]
                try:
                    return ast.literal_eval(expr)
                except Exception:
                    if expr in ("True", "False"):
                        return expr == "True"
                    return expr.strip('\"\'')

            dados['pasta_raiz'] = _get_val('PASTA_RAIZ') or ''
            dados['nome_mapa'] = _get_val('NOME_MAPA') or ''
            dados['pasta_copia'] = _get_val('PASTA_COPIA') or ''
            dados['sufixo_copia'] = _get_val('SUFIXO_COPIA') or ''
            dados['pasta_xlsx'] = _get_val('PASTA_XLSX') or ''
            dados['busca_recursiva'] = _get_val('BUSCA_RECURSIVA') if _get_val('BUSCA_RECURSIVA') is not None else True
            dados['project_visivel'] = _get_val('PROJECT_VISIVEL') if _get_val('PROJECT_VISIVEL') is not None else True

            col_match = re.search(r'COLUNAS\s*=\s*(\[[\s\S]*?\])', conteudo, re.MULTILINE)
            colunas = []
            if col_match:
                try:
                    col_list = ast.literal_eval(col_match.group(1))
                    colunas = [tuple(col) for col in col_list if isinstance(col, (list, tuple)) and len(col) == 2]
                except Exception:
                    colunas = []

            dados['colunas'] = colunas
            return dados
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar config: {e}")
            return {}
    
    def _salvar_config(self):
        """Salva todas as configurações no config.py"""
        try:
            self._atualizar_config_dados()
            
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            
            def format_path(path):
                """Formata caminho como raw string"""
                escaped = path.replace('"', '\\"')
                return f'r"{escaped}"'
            
            conteudo = re.sub(
                r'PASTA_RAIZ\s*=\s*.+',
                lambda m, rep=(f'PASTA_RAIZ = {format_path(self.config_dados["pasta_raiz"])}'): rep,
                conteudo,
                count=1
            )
            conteudo = re.sub(
                r'PASTA_COPIA\s*=\s*.+',
                lambda m, rep=(f'PASTA_COPIA = {format_path(self.config_dados["pasta_copia"])}'): rep,
                conteudo,
                count=1
            )
            conteudo = re.sub(
                r'PASTA_XLSX\s*=\s*.+',
                lambda m, rep=(f'PASTA_XLSX = {format_path(self.config_dados["pasta_xlsx"])}'): rep,
                conteudo,
                count=1
            )
            
            conteudo = re.sub(
                r'NOME_MAPA\s*=\s*.+',
                lambda m, rep=(f'NOME_MAPA = {repr(self.config_dados["nome_mapa"])}'): rep,
                conteudo,
                count=1
            )
            conteudo = re.sub(
                r'SUFIXO_COPIA\s*=\s*.+',
                lambda m, rep=(f'SUFIXO_COPIA = {repr(self.config_dados["sufixo_copia"])}'): rep,
                conteudo,
                count=1
            )
            
            conteudo = re.sub(
                r'BUSCA_RECURSIVA\s*=\s*(True|False)',
                lambda m, rep=(f'BUSCA_RECURSIVA = {str(self.config_dados["busca_recursiva"])}'): rep,
                conteudo,
                count=1
            )
            conteudo = re.sub(
                r'PROJECT_VISIVEL\s*=\s*(True|False)',
                lambda m, rep=(f'PROJECT_VISIVEL = {str(self.config_dados["project_visivel"])}'): rep,
                conteudo,
                count=1
            )
            
            novo_colunas = "COLUNAS = [\n"

            def format_literal(s: str) -> str:
                if "\\" in s:
                    escaped = s.replace('"', '\\"')
                    return f'r"{escaped}"'
                return repr(s)

            for nome, atributo in self.colunas:
                novo_colunas += f'    ({format_literal(nome)}, {format_literal(atributo)}),\n'

            novo_colunas += "]"
            
            inicio = conteudo.find('COLUNAS = [')
            if inicio != -1:
                fim = conteudo.find(']', inicio) + 1
                conteudo = conteudo[:inicio] + novo_colunas + conteudo[fim:]
            
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                f.write(conteudo)
            
            messagebox.showinfo("Sucesso", "Todas as configurações foram salvas!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {e}")
    
    def _criar_widgets(self):
        """Cria a interface com abas"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self._criar_aba_geral()
        
        self._criar_aba_colunas()
        
        frame_final = ttk.Frame(self.root)
        frame_final.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(frame_final, text="💾 Salvar Tudo", 
              command=self._salvar_config).pack(side=tk.RIGHT, padx=5)
        ttk.Button(frame_final, text="Fechar", 
              command=self.root.destroy).pack(side=tk.RIGHT, padx=5)
    
    def _criar_aba_geral(self):
        """Cria a aba de configurações gerais"""
        frame = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(frame, text="⚙️ Configurações Gerais")
        
        canvas = tk.Canvas(frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        ttk.Label(scrollable_frame, text="Pasta Raiz (origem .mpp):", 
                 font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=10)
        self.entry_pasta_raiz = ttk.Entry(scrollable_frame, width=50)
        self.entry_pasta_raiz.grid(row=0, column=1, padx=10)
        ttk.Button(scrollable_frame, text="📁 Procurar", 
                  command=lambda: self._selecionar_pasta('pasta_raiz')).grid(row=0, column=2)
        
        ttk.Label(scrollable_frame, text="Nome do Mapa:", 
                 font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.entry_nome_mapa = ttk.Entry(scrollable_frame, width=50)
        self.entry_nome_mapa.grid(row=1, column=1, padx=10)
        
        ttk.Label(scrollable_frame, text="Pasta de Cópias (processamento):", 
                 font=("Arial", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=10)
        self.entry_pasta_copia = ttk.Entry(scrollable_frame, width=50)
        self.entry_pasta_copia.grid(row=2, column=1, padx=10)
        ttk.Button(scrollable_frame, text="📁 Procurar", 
                  command=lambda: self._selecionar_pasta('pasta_copia')).grid(row=2, column=2)
        
        ttk.Label(scrollable_frame, text="Sufixo da Cópia:", 
                 font=("Arial", 10, "bold")).grid(row=3, column=0, sticky=tk.W, pady=10)
        self.entry_sufixo_copia = ttk.Entry(scrollable_frame, width=50)
        self.entry_sufixo_copia.grid(row=3, column=1, padx=10)
        
        ttk.Label(scrollable_frame, text="Pasta de Exportação (XLSX):", 
                 font=("Arial", 10, "bold")).grid(row=4, column=0, sticky=tk.W, pady=10)
        self.entry_pasta_xlsx = ttk.Entry(scrollable_frame, width=50)
        self.entry_pasta_xlsx.grid(row=4, column=1, padx=10)
        ttk.Button(scrollable_frame, text="📁 Procurar", 
                  command=lambda: self._selecionar_pasta('pasta_xlsx')).grid(row=4, column=2)
        
        self.var_busca_recursiva = tk.BooleanVar()
        ttk.Label(scrollable_frame, text="Buscar em Subpastas:", 
                 font=("Arial", 10, "bold")).grid(row=5, column=0, sticky=tk.W, pady=10)
        ttk.Checkbutton(scrollable_frame, variable=self.var_busca_recursiva).grid(row=5, column=1, sticky=tk.W)
        
        self.var_project_visivel = tk.BooleanVar()
        ttk.Label(scrollable_frame, text="Mostrar Janela do Project:", 
                 font=("Arial", 10, "bold")).grid(row=6, column=0, sticky=tk.W, pady=10)
        ttk.Checkbutton(scrollable_frame, variable=self.var_project_visivel).grid(row=6, column=1, sticky=tk.W)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _criar_aba_colunas(self):
        """Cria a aba de gerenciamento de colunas"""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="📊 Colunas de Exportação")
        
        frame_lista = ttk.LabelFrame(frame, text="Colunas Configuradas", padding=10)
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=10)
        
        scrollbar = ttk.Scrollbar(frame_lista)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(frame_lista, yscrollcommand=scrollbar.set, 
                                 font=("Courier", 9), height=12)
        self.listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        self.listbox.bind('<<ListboxSelect>>', self._ao_selecionar)
        
        frame_edit = ttk.LabelFrame(frame, text="Adicionar/Editar Coluna", padding=10)
        frame_edit.pack(fill=tk.X, pady=10)
        
        ttk.Label(frame_edit, text="Nome (XLSX):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.entry_nome = ttk.Entry(frame_edit, width=40)
        self.entry_nome.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame_edit, text="Atributo (Project):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.entry_atributo = ttk.Entry(frame_edit, width=40)
        self.entry_atributo.grid(row=1, column=1, padx=5, pady=5)
        
        frame_botoes = ttk.Frame(frame_edit)
        frame_botoes.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(frame_botoes, text="➕ Adicionar", 
                  command=self._adicionar_coluna).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="✏️ Editar", 
                  command=self._editar_coluna).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="🗑️ Remover", 
                  command=self._remover_coluna).pack(side=tk.LEFT, padx=5)
    
    def _atualizar_campos(self):
        """Atualiza os campos com os dados carregados"""
        if not self.config_dados:
            return
        
        self.entry_pasta_raiz.delete(0, tk.END)
        self.entry_pasta_raiz.insert(0, self.config_dados.get('pasta_raiz', ''))
        
        self.entry_nome_mapa.delete(0, tk.END)
        self.entry_nome_mapa.insert(0, self.config_dados.get('nome_mapa', ''))
        
        self.entry_pasta_copia.delete(0, tk.END)
        self.entry_pasta_copia.insert(0, self.config_dados.get('pasta_copia', ''))
        
        self.entry_sufixo_copia.delete(0, tk.END)
        self.entry_sufixo_copia.insert(0, self.config_dados.get('sufixo_copia', ''))
        
        self.entry_pasta_xlsx.delete(0, tk.END)
        self.entry_pasta_xlsx.insert(0, self.config_dados.get('pasta_xlsx', ''))
        
        self.var_busca_recursiva.set(self.config_dados.get('busca_recursiva', True))
        self.var_project_visivel.set(self.config_dados.get('project_visivel', True))
        
        self._atualizar_lista_colunas()
    
    def _atualizar_lista_colunas(self):
        """Atualiza a listbox de colunas"""
        self.listbox.delete(0, tk.END)
        for i, (nome, atributo) in enumerate(self.colunas):
            self.listbox.insert(tk.END, f"{i+1:2d}. {nome:30s} ← {atributo}")
    
    def _atualizar_config_dados(self):
        """Atualiza config_dados com os valores dos campos"""
        self.config_dados['pasta_raiz'] = self.entry_pasta_raiz.get()
        self.config_dados['nome_mapa'] = self.entry_nome_mapa.get()
        self.config_dados['pasta_copia'] = self.entry_pasta_copia.get()
        self.config_dados['sufixo_copia'] = self.entry_sufixo_copia.get()
        self.config_dados['pasta_xlsx'] = self.entry_pasta_xlsx.get()
        self.config_dados['busca_recursiva'] = self.var_busca_recursiva.get()
        self.config_dados['project_visivel'] = self.var_project_visivel.get()
    
    def _ao_selecionar(self, event):
        """Callback quando seleciona uma coluna"""
        seleção = self.listbox.curselection()
        if seleção:
            self.selecionada = seleção[0]
            nome, atributo = self.colunas[self.selecionada]
            self.entry_nome.delete(0, tk.END)
            self.entry_nome.insert(0, nome)
            self.entry_atributo.delete(0, tk.END)
            self.entry_atributo.insert(0, atributo)
    
    def _selecionar_pasta(self, campo):
        """Abre diálogo para selecionar pasta"""
        pasta = filedialog.askdirectory(title=f"Selecionar {campo}")
        if pasta:
            self.config_dados[campo] = pasta
            self._atualizar_campos()
    
    def _adicionar_coluna(self):
        """Adiciona uma nova coluna"""
        nome = self.entry_nome.get().strip()
        atributo = self.entry_atributo.get().strip()
        
        if not nome or not atributo:
            messagebox.showwarning("Aviso", "Preencha nome e atributo!")
            return
        
        if any(col[0] == nome for col in self.colunas):
            messagebox.showwarning("Aviso", f"Coluna '{nome}' já existe!")
            return
        
        self.colunas.append((nome, atributo))
        self.entry_nome.delete(0, tk.END)
        self.entry_atributo.delete(0, tk.END)
        self._atualizar_lista_colunas()
        messagebox.showinfo("Sucesso", "Coluna adicionada!")
    
    def _editar_coluna(self):
        """Edita a coluna selecionada"""
        if self.selecionada is None:
            messagebox.showwarning("Aviso", "Selecione uma coluna para editar!")
            return
        
        nome = self.entry_nome.get().strip()
        atributo = self.entry_atributo.get().strip()
        
        if not nome or not atributo:
            messagebox.showwarning("Aviso", "Preencha nome e atributo!")
            return
        
        for i, (n, _) in enumerate(self.colunas):
            if i != self.selecionada and n == nome:
                messagebox.showwarning("Aviso", f"Coluna '{nome}' já existe!")
                return
        
        self.colunas[self.selecionada] = (nome, atributo)
        self._atualizar_lista_colunas()
        self.selecionada = None
        self.entry_nome.delete(0, tk.END)
        self.entry_atributo.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Coluna editada!")
    
    def _remover_coluna(self):
        """Remove a coluna selecionada"""
        if self.selecionada is None:
            messagebox.showwarning("Aviso", "Selecione uma coluna para remover!")
            return
        
        nome = self.colunas[self.selecionada][0]
        if messagebox.askyesno("Confirmação", f"Remover coluna '{nome}'?"):
            self.colunas.pop(self.selecionada)
            self._atualizar_lista_colunas()
            self.selecionada = None
            self.entry_nome.delete(0, tk.END)
            self.entry_atributo.delete(0, tk.END)
    
    def _antes_fechar(self):
        """Salva automaticamente antes de fechar"""
        if messagebox.askyesno("Salvar", "Deseja salvar as alterações?"):
            self._atualizar_config_dados()
            self._salvar_config()
        try:
            self.root.destroy()
        except Exception:
            try:
                self.root.quit()
            except Exception:
                pass


if __name__ == "__main__":
    root = tk.Tk()
    app = ConfiguradorCompleto(root)
    root.protocol("WM_DELETE_WINDOW", app._antes_fechar)
    root.mainloop()

