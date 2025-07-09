import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import ttkbootstrap as tb
import pandas as pd
import os
import re
import threading
import sys
import subprocess # Para abrir pastas

# --- Funções Auxiliares ---

def get_output_path():
    """Retorna o caminho desejado para salvar os arquivos."""
    # O caminho fixo que você deseja usar
    return r"D:\Meus Arquivos\Documents\resultados"

def remove_accents(text):
    """Remove acentos de uma string para padronização."""
    text = str(text) # Garante que seja string
    text = re.sub(r'[ÁÀÂÃÄ]', 'A', text)
    text = re.sub(r'[ÉÈÊË]', 'E', text)
    text = re.sub(r'[ÍÌÎÏ]', 'I', text)
    text = re.sub(r'[ÓÒÔÕÖ]', 'O', text)
    text = re.sub(r'[ÚÙÛÜ]', 'U', text)
    text = re.sub(r'[Ç]', 'C', text)
    return text

def open_folder(path):
    """Abre um diretório no Explorador de Arquivos (Windows)."""
    try:
        os.makedirs(path, exist_ok=True) # Garante que o diretório exista antes de tentar abrir
        # Usamos subprocess.Popen para maior compatibilidade
        if sys.platform == "win32":
            subprocess.Popen(['explorer', path])
        elif sys.platform == "darwin": # macOS
            subprocess.Popen(["open", path])
        else: # Linux
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Erro ao Abrir Pasta", f"Não foi possível abrir o diretório: {e}")

def column_letter_to_index(column_letter):
    """Converte uma letra de coluna Excel (A, B, C...) para o índice baseado em zero."""
    if not isinstance(column_letter, str) or not column_letter.isalpha() or len(column_letter) != 1:
        raise ValueError("A letra da coluna deve ser uma única letra alfabética.")
    
    # Converte para maiúscula para padronização
    column_letter = column_letter.upper()
    
    # 'A' é 0, 'B' é 1, ..., 'Z' é 25
    return ord(column_letter) - ord('A')

# --- Lógica de Processamento do Excel ---

def process_excel(excel_path, prova_nome, serie_selecionada, column_note_index, progress_bar, status_label, not_found_text_area, root):
    """
    Processa o arquivo Excel, compara os dados e gera o arquivo TXT.
    Executado em uma thread separada.
    """
    has_occurrences = False # Flag para indicar se houve múltiplas correspondências (ambiguidade)
    has_partial_matches = False # Flag para indicar se houve coincidências parciais
    not_found_alunos = [] # Lista para armazenar alunos não encontrados

    try:
        status_label.config(text="Carregando arquivo Excel...")
        progress_bar['value'] = 0
        root.update_idletasks() # Atualiza a GUI
        not_found_text_area.config(state='normal') # Habilita para escrita na thread
        not_found_text_area.delete(1.0, tk.END) # Limpa a área de alunos não encontrados
        not_found_text_area.config(state='disabled') # Desabilita novamente

        if not os.path.exists(excel_path):
            messagebox.showerror("Erro de Arquivo", f"Arquivo Excel não encontrado: '{excel_path}'. Por favor, verifique o caminho.")
            status_label.config(text="Erro: Arquivo Excel não encontrado.")
            return

        xls = pd.ExcelFile(excel_path, engine='openpyxl')

        try:
            df_serie = pd.read_excel(xls, sheet_name=serie_selecionada, header=None, skiprows=6)
        except ValueError:
            messagebox.showerror("Erro de Planilha", f"A planilha '{serie_selecionada}' não foi encontrada no arquivo Excel. Verifique o nome da planilha.")
            status_label.config(text="Erro: Planilha da série não encontrada.")
            return

        try:
            df_lista_alunos = pd.read_excel(xls, sheet_name='Lista de Alunos', header=None)
        except ValueError:
            messagebox.showerror("Erro de Planilha", "A planilha 'Lista de Alunos' não foi encontrada no arquivo Excel. Verifique o nome da planilha.")
            status_label.config(text="Erro: Planilha 'Lista de Alunos' não encontrada.")
            return

        status_label.config(text="Pré-processando dados...")
        progress_bar['value'] = 10
        root.update_idletasks()

        # 2. Pré-processar DataFrames
        # Agora usando column_note_index
        df_serie = df_serie[[0, column_note_index]].copy()
        df_serie.columns = ['NomeAlunoSerie', 'Nota']
        df_serie['NomeAlunoSerie'] = df_serie['NomeAlunoSerie'].astype(str).str.strip().str.upper()
        df_serie['Nota'] = df_serie['Nota'].astype(str).str.replace(',', '.', regex=False)
        df_serie['Nota'] = pd.to_numeric(df_serie['Nota'], errors='coerce')
        df_serie.dropna(subset=['Nota'], inplace=True)

        df_lista_alunos = df_lista_alunos[[0, 1]].copy()
        df_lista_alunos.columns = ['Matricula', 'NomeCompletoLista']
        df_lista_alunos['NomeCompletoLista'] = df_lista_alunos['NomeCompletoLista'].astype(str).str.strip().str.upper()
        df_lista_alunos['Matricula'] = df_lista_alunos['Matricula'].astype(str).str.strip()

        progress_bar['value'] = 30
        root.update_idletasks()

        # 3. Comparar e gerar os arquivos TXT
        output_file_name_main = f"{serie_selecionada} - {prova_nome}.txt"
        output_file_name_ambiguities = f"ocorrencias {serie_selecionada} - {prova_nome}.txt" # Múltiplas matches, ambiguidade
        output_file_name_partial = f"Coincidencias parciais {serie_selecionada} - {prova_nome}.txt" # Novo: Coincidências parciais
        output_file_name_not_found = "alunos_nao_encontrados.txt" # Novo: Alunos não encontrados

        destination_path = get_output_path()
        main_output_file_path = os.path.join(destination_path, output_file_name_main)
        ambiguities_output_file_path = os.path.join(destination_path, output_file_name_ambiguities)
        partial_output_file_path = os.path.join(destination_path, output_file_name_partial)
        not_found_output_file_path = os.path.join(destination_path, output_file_name_not_found)

        matched_alunos = []
        ambiguities_list = [] # Para armazenar múltiplas correspondências (ambiguidades)
        partial_matches_log = [] # Para armazenar logs de coincidências parciais
        
        processed_names_serie = set() # Para garantir nomes únicos da planilha de série

        total_alunos_serie = len(df_serie)
        if total_alunos_serie == 0:
            messagebox.showinfo("Processamento", f"Nenhum aluno com nota válida encontrado na planilha '{serie_selecionada}'.")
            status_label.config(text="Nenhum aluno processado.")
            return

        # Pre-calcula a coluna padronizada na df_lista_alunos uma vez
        df_lista_alunos['NomeCompletoLista_STD'] = df_lista_alunos['NomeCompletoLista'].apply(remove_accents).str.strip().str.upper()

        for index, row_serie in df_serie.iterrows():
            nome_aluno_serie_original = row_serie['NomeAlunoSerie']
            nota_serie = row_serie['Nota']

            if nome_aluno_serie_original in processed_names_serie:
                continue
            processed_names_serie.add(nome_aluno_serie_original)

            nome_aluno_serie_std = remove_accents(nome_aluno_serie_original).strip().upper()

            if not nome_aluno_serie_std: # Evita erro se o nome estiver vazio após limpeza
                print(f"Alerta: Nome vazio ou inválido na série: '{nome_aluno_serie_original}'. Ignorado.")
                continue

            pattern_parts = [re.escape(part) for part in re.split(r'\s+', nome_aluno_serie_std) if part]
            regex_pattern_str = r".*".join(pattern_parts)
            regex_pattern = re.compile(f"^{regex_pattern_str}.*$", re.IGNORECASE)

            potential_matches = df_lista_alunos[df_lista_alunos['NomeCompletoLista_STD'].str.contains(regex_pattern, na=False)]

            found_match_for_student = False

            if len(potential_matches) == 1:
                matched_name_lista_std = potential_matches.iloc[0]['NomeCompletoLista_STD']
                matricula = potential_matches.iloc[0]['Matricula']

                if matched_name_lista_std == nome_aluno_serie_std:
                    nota_formatada = f"{nota_serie:.1f}".replace('.',',')
                    matched_alunos.append(f"{matricula}\t{nota_formatada}")
                    found_match_for_student = True
                elif nome_aluno_serie_std.startswith(matched_name_lista_std) and len(nome_aluno_serie_std) > len(matched_name_lista_std):
                    matched_alunos.append(f"{matricula}\t{nota_serie:.1f}")
                    partial_matches_log.append(
                        f"Aluno da Série: '{nome_aluno_serie_original}' (Nota: {nota_serie:.1f})\n"
                        f"  Coincidência Parcial com Lista de Alunos:\n"
                        f"  - Matrícula: {matricula}, Nome Completo: '{potential_matches.iloc[0]['NomeCompletoLista']}'\n"
                        f"----------------------------------------\n"
                    )
                    has_partial_matches = True
                    found_match_for_student = True
                # Else: single regex match but not exact or prefix, treated as not found for now
            elif len(potential_matches) > 1:
                has_occurrences = True
                ambiguity_info = f"Aluno da Série: '{nome_aluno_serie_original}' (Nota: {nota_serie:.1f})\n"
                ambiguity_info += "Possíveis correspondências (Ambíguas) na Lista de Alunos:\n"
                for idx, row_match in potential_matches.iterrows():
                    ambiguity_info += f"- Matrícula: {row_match['Matricula']}, Nome: '{row_match['NomeCompletoLista']}'\n"
                ambiguity_info += "----------------------------------------\n"
                ambiguities_list.append(ambiguity_info)
            
            # Adicionar aluno à lista de não encontrados se nenhuma das condições acima for satisfeita
            if not found_match_for_student and len(potential_matches) == 0:
                not_found_alunos.append(nome_aluno_serie_original)


            # Atualiza barra de progresso
            progress = (index + 1) / total_alunos_serie * 100
            progress_bar['value'] = 30 + (progress * 0.6)
            status_label.config(text=f"Processando: {int(progress)}% - {nome_aluno_serie_original}")
            root.update_idletasks()

        # Remove a coluna temporária usada para padronização
        df_lista_alunos.drop(columns=['NomeCompletoLista_STD'], inplace=True)

        progress_bar['value'] = 90
        status_label.config(text="Gerando arquivos de saída...")
        root.update_idletasks()

        # Garante que o diretório de destino exista
        os.makedirs(destination_path, exist_ok=True)

        # 4. Escrever no arquivo TXT principal
        if matched_alunos:
            try:
                with open(main_output_file_path, 'w', encoding='utf-8') as f:
                    for line in matched_alunos:
                        f.write(line + '\n')
            except IOError as e:
                messagebox.showerror("Erro de Escrita", f"Não foi possível salvar o arquivo principal '{output_file_name_main}'. Erro: {e}")
                status_label.config(text="Erro de escrita principal.")
                progress_bar['value'] = 0
                return
        else:
            messagebox.showwarning("Aviso", "Nenhum aluno correspondente único encontrado para o arquivo TXT principal.")


        # 5. Escrever no arquivo de ambigüidades (se houver)
        if ambiguities_list:
            try:
                with open(ambiguities_output_file_path, 'w', encoding='utf-8') as f:
                    for entry in ambiguities_list:
                        f.write(entry + '\n')
            except IOError as e:
                messagebox.showerror("Erro de Escrita", f"Não foi possível salvar o arquivo de ambigüidades '{output_file_name_ambiguities}'. Erro: {e}")
                status_label.config(text="Erro de escrita de ambigüidades.")
                progress_bar['value'] = 0
                return

        # 6. Escrever no arquivo de coincidências parciais (se houver)
        if partial_matches_log:
            try:
                with open(partial_output_file_path, 'a', encoding='utf-8') as f:
                    if os.path.getsize(partial_output_file_path) == 0:
                        f.write(f"--- Coincidências Parciais em {serie_selecionada} - {prova_nome} ---\n\n")
                    else:
                        f.write(f"\n--- Novas Coincidências em {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n\n")
                    for entry in partial_matches_log:
                        f.write(entry + '\n')
            except IOError as e:
                messagebox.showerror("Erro de Escrita", f"Não foi possível salvar o arquivo de coincidências parciais '{output_file_name_partial}'. Erro: {e}")
                status_label.config(text="Erro de escrita de coincidências parciais.")
                progress_bar['value'] = 0
                return

        # 7. Escrever no arquivo de alunos não encontrados (se houver)
        if not_found_alunos:
            try:
                with open(not_found_output_file_path, 'a', encoding='utf-8') as f:
                    if os.path.getsize(not_found_output_file_path) == 0:
                        f.write(f"--- Alunos Não Encontrados em {serie_selecionada} - {prova_nome} ---\n\n")
                    else:
                        f.write(f"\n--- Novas Entradas em {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n\n")
                    for aluno in not_found_alunos:
                        f.write(aluno + '\n')
            except IOError as e:
                messagebox.showerror("Erro de Escrita", f"Não foi possível salvar o arquivo de alunos não encontrados '{output_file_name_not_found}'. Erro: {e}")
                status_label.config(text="Erro de escrita de alunos não encontrados.")
                progress_bar['value'] = 0
                return
            
            # Exibir alunos não encontrados na área de texto da GUI
            not_found_text_area.config(state='normal')
            not_found_text_area.insert(tk.END, "\nAlunos Não Encontrados:\n")
            not_found_text_area.insert(tk.END, "-------------------------\n")
            for aluno in not_found_alunos:
                not_found_text_area.insert(tk.END, aluno + "\n")
            not_found_text_area.see(tk.END) # Rola para o final
            not_found_text_area.config(state='disabled')
        else:
            not_found_text_area.config(state='normal')
            not_found_text_area.insert(tk.END, "\nNenhum aluno não encontrado neste processamento.\n")
            not_found_text_area.see(tk.END)
            not_found_text_area.config(state='disabled')


        # Mensagem de conclusão modificada
        final_message = "Processamento finalizado!\n"
        final_message += f"Arquivo principal '{output_file_name_main}' gerado em '{destination_path}'.\n"

        if has_partial_matches:
            final_message += f"\nATENÇÃO: Houve coincidências parciais de nomes. Verifique o arquivo '{output_file_name_partial}' para revisão."
        if has_occurrences:
            final_message += f"\nATENÇÃO: Houve coincidências ambíguas de nomes. Verifique o arquivo '{output_file_name_ambiguities}' para revisão manual."
        if not_found_alunos:
            final_message += f"\nATENÇÃO: Alunos não encontrados. Verifique a lista abaixo ou o arquivo '{output_file_name_not_found}' para revisão."

        if not has_partial_matches and not has_occurrences and not not_found_alunos:
             final_message += "\nNenhuma ocorrência ou divergência de nomes foi encontrada."

        messagebox.showinfo("Concluído", final_message)
        status_label.config(text="Concluído!")
        progress_bar['value'] = 100

    except FileNotFoundError as e:
        messagebox.showerror("Erro de Arquivo", f"Um arquivo necessário não foi encontrado: {e}. Verifique se o arquivo Excel existe.")
        status_label.config(text="Erro: Arquivo não encontrado.")
        progress_bar['value'] = 0
    except pd.errors.EmptyDataError:
        messagebox.showerror("Erro de Dados", "O arquivo Excel está vazio ou não contém dados na planilha selecionada.")
        status_label.config(text="Erro: Dados vazios no Excel.")
        progress_bar['value'] = 0
    except Exception as e:
        messagebox.showerror("Erro Geral", f"Ocorreu um erro inesperado: {e}")
        status_label.config(text="Erro: " + str(e))
        progress_bar['value'] = 0
    finally:
        root.update_idletasks()


# --- Interface Gráfica (Tkinter/ttkbootstrap) ---

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Processador de Notas de Alunos")
        self.root.geometry("700x650") # Aumenta a altura da janela
        self.root.resizable(False, False)

        self.excel_file_path = tk.StringVar()
        self.prova_nome = tk.StringVar()
        self.serie_selecionada = tk.StringVar()
        self.series_opcoes = ["1ª Série", "2ª Série", "3ª Série"]
        self.coluna_nota = tk.StringVar() # Novo StringVar para a coluna da nota

        self.create_widgets()
        # Inicializa o monitoramento para validação
        self.coluna_nota.trace_add("write", self._validate_column_input)
        self.excel_file_path.trace_add("write", self._check_all_inputs_valid)
        self.prova_nome.trace_add("write", self._check_all_inputs_valid)
        self.serie_selecionada.trace_add("write", self._check_all_inputs_valid)
        
        # Chama a validação inicial para desabilitar o botão se os campos estiverem vazios
        self._check_all_inputs_valid()


    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Frame para título e botões de ação (Ajuda/Abrir Pasta)
        top_bar_frame = ttk.Frame(main_frame)
        top_bar_frame.pack(pady=5, fill=tk.X)

        title_label = ttk.Label(top_bar_frame, text="Processador de Notas Excel", font=("Helvetica", 16, "bold"))
        title_label.pack(side=tk.LEFT, padx=(0, 20)) # Ajusta padding para o título

        # Botão de Ajuda
        help_button = ttk.Button(top_bar_frame, text="Ajuda", command=self.show_help, bootstyle="info-outline")
        help_button.pack(side=tk.RIGHT, padx=(5, 0))


        # Campos de entrada
        prova_frame = ttk.Frame(main_frame)
        prova_frame.pack(pady=5, fill=tk.X)
        ttk.Label(prova_frame, text="Nome da Prova:").pack(side=tk.LEFT, padx=(0, 10))
        self.prova_entry = ttk.Entry(prova_frame, textvariable=self.prova_nome, width=40)
        self.prova_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        serie_frame = ttk.Frame(main_frame)
        serie_frame.pack(pady=5, fill=tk.X)
        ttk.Label(serie_frame, text="Série:").pack(side=tk.LEFT, padx=(0, 10))
        self.serie_dropdown = ttk.OptionMenu(serie_frame, self.serie_selecionada, self.series_opcoes[0], *self.series_opcoes)
        self.serie_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Novo campo para a Coluna da Nota
        coluna_nota_frame = ttk.Frame(main_frame)
        coluna_nota_frame.pack(pady=5, fill=tk.X)
        ttk.Label(coluna_nota_frame, text="Coluna da Nota (ex: 'N' ou 'P'):").pack(side=tk.LEFT, padx=(0, 10))
        self.coluna_nota_entry = ttk.Entry(coluna_nota_frame, textvariable=self.coluna_nota, width=5)
        self.coluna_nota_entry.pack(side=tk.LEFT, padx=(0, 10))
        # Definir estilo inicial para o campo de input da coluna
        self.coluna_nota_entry.config(bootstyle="default") # Garante um estilo padrão


        excel_frame = ttk.Frame(main_frame)
        excel_frame.pack(pady=10, fill=tk.X)

        ttk.Label(excel_frame, text="Arquivo Excel:").pack(side=tk.LEFT, padx=(0, 10))
        self.excel_path_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_path, width=50, state='readonly')
        self.excel_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.browse_button = ttk.Button(excel_frame, text="Procurar", command=self.browse_excel_file)
        self.browse_button.pack(side=tk.LEFT, padx=(10, 0))

        self.start_button = ttk.Button(main_frame, text="Iniciar Processamento", command=self.start_processing_thread, bootstyle="primary")
        self.start_button.pack(pady=20)

        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate', length=400)
        self.progress_bar.pack(pady=10)

        self.status_label = ttk.Label(main_frame, text="Pronto para processar.", bootstyle="info")
        self.status_label.pack(pady=5)
        
        # --- Nova área para exibir alunos não encontrados ---
        ttk.Label(main_frame, text="Alunos Não Encontrados Recentes:", font=("Helvetica", 10, "bold")).pack(pady=(15, 5))
        self.not_found_text_area = scrolledtext.ScrolledText(main_frame, width=70, height=8, wrap=tk.WORD, font=("Consolas", 9))
        self.not_found_text_area.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        self.not_found_text_area.config(state='disabled') # Torna a caixa de texto somente leitura

    def _validate_column_input(self, *args):
        """Valida a entrada da coluna da nota (deve ser uma única letra)."""
        value = self.coluna_nota.get().strip()
        is_valid = bool(value and value.isalpha() and len(value) == 1)

        if is_valid:
            self.coluna_nota_entry.config(bootstyle="default") # Volta ao estilo normal
        else:
            self.coluna_nota_entry.config(bootstyle="danger") # Estilo de erro (vermelho)
        
        self._check_all_inputs_valid() # Re-verifica o botão de processamento

    def _check_all_inputs_valid(self, *args):
        """Verifica se todos os inputs necessários são válidos para habilitar o botão de processamento."""
        excel_ok = bool(self.excel_file_path.get())
        prova_ok = bool(self.prova_nome.get().strip())
        serie_ok = bool(self.serie_selecionada.get())
        coluna_ok = bool(self.coluna_nota.get().strip() and self.coluna_nota.get().strip().isalpha() and len(self.coluna_nota.get().strip()) == 1)

        if excel_ok and prova_ok and serie_ok and coluna_ok:
            self.start_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.DISABLED)

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_file_path.set(file_path)
            self.status_label.config(text="Arquivo selecionado.")
            self.progress_bar['value'] = 0
            self.not_found_text_area.config(state='normal')
            self.not_found_text_area.delete(1.0, tk.END) # Limpa área ao selecionar novo arquivo
            self.not_found_text_area.config(state='disabled')
            self._check_all_inputs_valid() # Re-verifica o botão

    def start_processing_thread(self):
        excel_path = self.excel_file_path.get()
        prova_nome = self.prova_nome.get().strip()
        serie_selecionada = self.serie_selecionada.get()
        coluna_nota_letra = self.coluna_nota.get().strip()

        # Re-validação final (redundante, mas seguro)
        if not (excel_path and prova_nome and serie_selecionada and coluna_nota_letra and coluna_nota_letra.isalpha() and len(coluna_nota_letra) == 1):
            messagebox.showwarning("Entrada Inválida", "Por favor, preencha e valide todos os campos corretamente.")
            return
        
        try:
            coluna_nota_index = column_letter_to_index(coluna_nota_letra)
        except ValueError as e:
            messagebox.showerror("Erro de Coluna", str(e))
            self.status_label.config(text="Erro: Coluna da nota inválida.")
            return

        self.start_button.config(state=tk.DISABLED)
        self.status_label.config(text="Iniciando processamento...")
        self.progress_bar['value'] = 0
        self.not_found_text_area.config(state='normal') # Habilita para escrita na thread

        process_thread = threading.Thread(
            target=process_excel,
            args=(excel_path, prova_nome, serie_selecionada, coluna_nota_index,
                  self.progress_bar, self.status_label, self.not_found_text_area, self.root)
        )
        process_thread.start()

        self.check_thread(process_thread)

    def check_thread(self, thread):
        if thread.is_alive():
            self.root.after(100, lambda: self.check_thread(thread))
        else:
            self.start_button.config(state=tk.NORMAL)
            self.not_found_text_area.config(state='disabled') # Desabilita novamente após a thread terminar

    def show_help(self):
        """Exibe a janela de ajuda com explicações e botão para abrir pasta."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Ajuda - Processador de Notas")
        help_window.geometry("500x380") # Aumenta um pouco a altura para o novo item
        help_window.resizable(False, False)
        help_window.transient(self.root) # Torna a janela de ajuda filha da principal
        help_window.grab_set() # Bloqueia interações com a janela principal

        help_frame = ttk.Frame(help_window, padding=20)
        help_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(help_frame, text="Como Usar o Processador de Notas", font=("Helvetica", 14, "bold")).pack(pady=10)

        help_text = (
            "1. Nome da Prova: Digite um nome para a prova. Ele será usado para nomear os arquivos de saída (ex: 'Avaliação Final').\n\n"
            "2. Série: Selecione a série correspondente à planilha de notas no arquivo Excel (ex: '1ª Série').\n\n"
            "3. Coluna da Nota: Digite a letra da coluna (ex: 'N' ou 'P') onde as notas estão localizadas na planilha da série. Apenas uma letra é aceita.\n\n"
            "4. Arquivo Excel: Clique em 'Procurar' para selecionar o arquivo Excel (.xlsx ou .xls) com as notas e a lista de alunos.\n\n"
            "5. Iniciar Processamento: Clique neste botão para iniciar a análise e a geração dos arquivos TXT.\n\n"
            "Arquivos Gerados:\n"
            "- {Série} - {Prova}.txt: Contém Matrícula e Nota dos alunos encontrados.\n"
            "- ocorrencias {Série} - {Prova}.txt: Lista alunos com nomes ambíguos que precisam de revisão manual.\n"
            "- Coincidencias parciais {Série} - {Prova}.txt: Lista alunos com nomes parcialmente correspondentes (ex: truncados).\n"
            "- alunos_nao_encontrados.txt: Lista alunos da planilha de série que não foram encontrados na 'Lista de Alunos'."
        )
        help_display = scrolledtext.ScrolledText(help_frame, width=60, height=14, wrap=tk.WORD, font=("TkDefaultFont", 9))
        help_display.insert(tk.END, help_text)
        help_display.config(state='disabled') # Torna a caixa de texto somente leitura
        help_display.pack(pady=10, fill=tk.BOTH, expand=True)

        open_folder_button = ttk.Button(help_frame, text="Abrir Pasta de Saída", command=lambda: open_folder(get_output_path()), bootstyle="success")
        open_folder_button.pack(pady=10)


# --- Inicialização da Aplicação ---
if __name__ == "__main__":
    app_root = tb.Window(themename="darkly") # Escolha um tema ttkbootstrap
    ExcelProcessorApp(app_root)
    app_root.mainloop()