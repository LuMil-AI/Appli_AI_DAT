#Appli Analyseur DAT
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import font as tkfont
import csv
import os
import re

try:
    import pandas as pd
except ImportError:
    pd = None
    
# =================================================================================
# NOUVELLE CLASSE : TABLEAU AUTONOME POUR LA COMPARAISON
# =================================================================================
class TableWidget(tk.Frame):
    """
    Composant réutilisable gérant un tableau (Treeview) avec chargement type "Ouvrir Autre .DAT", 
    édition, sauvegarde et filtrage des colonnes.
    """
    def __init__(self, parent, title="Tableau", accent_color="#3498db"):
        super().__init__(parent, bg="#ecf0f1")
        self.accent_color = accent_color
        self.data = []
        self.headers = []
        self.visible_columns = []
        self.filtered_indices = []
        self.undo_stack = []
        self.file_path = None
        self.comparison_window = None
        
        # --- Toolbar ---
        toolbar = tk.Frame(self, bg="#dfe6e9", height=40)
        toolbar.pack(fill='x', side='top')
        
        tk.Label(toolbar, text=title, bg="#dfe6e9", font=("Segoe UI", 10, "bold")).pack(side='left', padx=10)
        
        btn_style = {"bg": accent_color, "fg": "white", "relief": "flat", "padx": 10, "pady": 2, "font": ("Segoe UI", 9)}
        
        tk.Button(toolbar, text="Charger", command=self.load_file, **btn_style).pack(side='left', padx=2)
        tk.Button(toolbar, text="Enregistrer", command=self.save_file, **btn_style).pack(side='left', padx=2)
        tk.Button(toolbar, text="Colonnes", command=self.select_columns, **btn_style).pack(side='left', padx=2)
        
        # --- Treeview ---
        self.tree_frame = tk.Frame(self)
        self.tree_frame.pack(fill='both', expand=True)
        
        self.h_scroll = ttk.Scrollbar(self.tree_frame, orient="horizontal")
        self.v_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical")
        
        self.tree = ttk.Treeview(self.tree_frame, 
                                 yscrollcommand=self.v_scroll.set, 
                                 xscrollcommand=self.h_scroll.set,
                                 selectmode="extended",
                                 show="headings")
        
        self.h_scroll.config(command=self.tree.xview)
        self.v_scroll.config(command=self.tree.yview)
        
        self.h_scroll.pack(side='bottom', fill='x')
        self.v_scroll.pack(side='right', fill='y')
        self.tree.pack(side='left', fill='both', expand=True)
        
        self.tree.bind('<Double-1>', self.edit_cell)
        self.tree.bind('<Button-3>', self.show_context_menu) # Windows / Linux
        self.tree.bind('<Button-2>', self.show_context_menu) # MacOS (parfois)
        self.tree.bind('<Control-z>', self.undo)

        # État pour la recherche "Suivant"
        self.last_search_index = -1

    def load_file(self):
        filetypes = [
            ("Tous supportés", "*.dat *.csv *.xlsx *.xls"),
            ("Fichiers DAT", "*.dat"),
            ("Fichiers CSV", "*.csv"),
            ("Fichiers Excel", "*.xlsx *.xls")
        ]
        path = filedialog.askopenfilename(parent=self, filetypes=filetypes)
        if not path:
            return
            
        self.file_path = path
        filename = os.path.basename(path).lower() # On récupère le nom du fichier en minuscule
        ext = os.path.splitext(path)[1].lower()
        
        try:
            self.data = []
            self.headers = []
            
            # --- Lecture du fichier (Excel ou CSV/DAT) ---
            if ext in ['.xlsx', '.xls']:
                if pd is None:
                    messagebox.showerror("Erreur", "Pandas n'est pas installé. Impossible d'ouvrir Excel.", parent=self)
                    return
                df = pd.read_excel(path, dtype=str).fillna("")
                self.headers = list(df.columns)
                self.data = df.values.tolist()
            else:
                with open(path, 'r', encoding='latin-1', errors='replace') as f:
                    reader = csv.reader(f, delimiter=';') 
                    try:
                        sample = f.read(1024)
                        f.seek(0)
                        if ',' in sample and ';' not in sample:
                            reader = csv.reader(f, delimiter=',')
                        else:
                            f.seek(0)
                            reader = csv.reader(f, delimiter=';')
                    except:
                        f.seek(0)
                        reader = csv.reader(f, delimiter=';')

                    self.data = list(reader)

            # --- DÉTECTION AUTOMATIQUE DES HEADERS ---
            # On vérifie si le nom du fichier contient un mot clé connu (varexp, comm, event, etc.)
            # Et on va chercher la constante correspondante dans la classe DatEditor
            detected_headers = []
            
            # Mapping entre mot-clé dans le fichier et nom de la variable dans DatEditor
            header_map = {
                "varexp": "VAREXP_DEFAULT_HEADERS",
                "event": "EVENT_DEFAULT_HEADERS",
                "comm": "COMM_DEFAULT_HEADERS",
                "vartreat": "VARTREAT_DEFAULT_HEADERS",
                "exprv": "EXPRV_DEFAULT_HEADERS",
                "cyclic": "CYCLIC_DEFAULT_HEADERS"
            }

            for key, attr_name in header_map.items():
                if key in filename:
                    # On essaie de récupérer la liste depuis la classe DatEditor
                    # (On utilise getattr pour éviter que ça plante si DatEditor n'est pas encore défini ou si l'attribut manque)
                    try:
                        detected_headers = getattr(DatEditor, attr_name, [])
                    except NameError:
                        pass # DatEditor pas encore défini ? Peu probable au runtime.
                    break
            
            # Si on a trouvé des headers spécifiques et que le fichier n'est pas un Excel (qui a ses propres headers)
            if detected_headers and ext not in ['.xlsx', '.xls']:
                self.headers = detected_headers.copy()
            
            # --- Ajustement des données aux headers ---
            if self.data:
                # Si on a défini des headers (automatiques ou Excel), on s'assure que les lignes font la bonne taille
                if self.headers:
                    target_len = len(self.headers)
                    # Si les headers auto sont plus longs que les données, on étend les données
                    # Si les données sont plus longues, on étend les headers (cas générique)
                    max_data_len = max(len(row) for row in self.data)
                    
                    final_len = max(target_len, max_data_len)
                    
                    # Si les données dépassent les headers prévus, on complète les headers
                    if final_len > len(self.headers):
                         for i in range(len(self.headers), final_len):
                             self.headers.append(f"Col_{i+1}")
                    
                    # On pad les données
                    for row in self.data:
                        if len(row) < final_len:
                            row.extend([""] * (final_len - len(row)))
                            
                else:
                    # Cas générique sans headers détectés
                    max_cols = max(len(row) for row in self.data)
                    for row in self.data:
                        if len(row) < max_cols:
                            row.extend([""] * (max_cols - len(row)))
                    self.headers = [f"Col_{i+1}" for i in range(max_cols)]
            else:
                 if not self.headers: self.headers = []

            self.visible_columns = self.headers.copy()
            self.filtered_indices = list(range(len(self.data)))
            self.refresh_tree()
            self.last_search_index = -1
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du chargement : {e}", parent=self)

    def save_file(self):
        if not self.data:
            return
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".csv", 
                                          filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("DAT", "*.dat")])
        if not path:
            return
            
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ['.xlsx', '.xls']:
                if pd is None:
                    messagebox.showerror("Erreur", "Pandas requis pour Excel.", parent=self)
                    return
                df = pd.DataFrame(self.data, columns=self.headers)
                df.to_excel(path, index=False)
            else:
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(self.headers)
                    writer.writerows(self.data)
            messagebox.showinfo("Succès", "Fichier enregistré.", parent=self)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur sauvegarde : {e}", parent=self)

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        display_cols = [c for c in self.headers if c in self.visible_columns]
        self.tree["columns"] = display_cols
        
        for col in display_cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, stretch=True)
            
        for i in self.filtered_indices:
            row = self.data[i]
            values = []
            for col in display_cols:
                try:
                    idx = self.headers.index(col)
                    val = row[idx] if idx < len(row) else ""
                    values.append(val)
                except ValueError:
                    values.append("")
            self.tree.insert("", "end", iid=str(i), values=values)

    def edit_cell(self, event):
        item_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item_id or not column: return

        row_index = int(item_id)
        col_num = int(column.replace('#', '')) - 1
        display_cols = [c for c in self.headers if c in self.visible_columns]
        if col_num < 0 or col_num >= len(display_cols): return
            
        col_name = display_cols[col_num]
        try: real_col_index = self.headers.index(col_name)
        except ValueError: return
        
        x, y, w, h = self.tree.bbox(item_id, column)
        current_val = self.data[row_index][real_col_index] if real_col_index < len(self.data[row_index]) else ""
        
        entry = tk.Entry(self.tree, font=("Segoe UI", 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, str(current_val))
        entry.focus()
        
        def save_edit(e):
            new_val = entry.get()
            self.data[row_index][real_col_index] = new_val
            self.tree.set(item_id, column, new_val)
            entry.destroy()
            
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: entry.destroy())

    def select_columns(self):
        if not self.headers: return
        top = tk.Toplevel(self)
        top.title("Colonnes")
        top.geometry("450x600")
        top.transient(self)
        
        # Toolbar boutons sélection
        btn_frame = tk.Frame(top)
        btn_frame.pack(fill='x', pady=5)
        
        vars = {}
        
        def set_all(value):
            for v in vars.values(): v.set(value)
            
        tk.Button(btn_frame, text="Tout sélectionner", command=lambda: set_all(True), bg="#95a5a6", fg="white", relief="flat").pack(side='left', padx=5)
        tk.Button(btn_frame, text="Tout désélectionner", command=lambda: set_all(False), bg="#95a5a6", fg="white", relief="flat").pack(side='left', padx=5)

        # Liste
        canvas = tk.Canvas(top)
        scrollbar = ttk.Scrollbar(top, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        for col in self.headers:
            v = tk.BooleanVar(value=col in self.visible_columns)
            chk = tk.Checkbutton(scrollable_frame, text=col, variable=v)
            chk.pack(anchor='w')
            vars[col] = v
            
        def apply():
            self.visible_columns = [c for c in self.headers if vars[c].get()]
            self.refresh_tree()
            top.destroy()
        
        # --- ICI LE BOUTON VALIDER DEMANDÉ ---
        tk.Button(top, text="Valider", command=apply, bg=self.accent_color, fg="white", font=("Segoe UI", 10, "bold"), pady=5).pack(fill='x', side='bottom')

    def find_next_and_replace(self, search_text, replace_text, do_replace=False):
        if not search_text: return 0
        search_lower = search_text.lower()
        start_idx = self.last_search_index + 1
        
        replaced = False
        if do_replace and self.last_search_index != -1 and self.last_search_index < len(self.data):
             row = self.data[self.last_search_index]
             changed = False
             for c_idx, cell in enumerate(row):
                 if search_lower in str(cell).lower():
                     pattern = re.compile(re.escape(search_text), re.IGNORECASE)
                     new_val = pattern.sub(replace_text, str(cell))
                     if new_val != str(cell):
                         self.data[self.last_search_index][c_idx] = new_val
                         changed = True
             if changed:
                 self.refresh_row(self.last_search_index)
                 replaced = True

        for i in range(start_idx, len(self.data)):
            row = self.data[i]
            row_str = " ".join([str(x).lower() for x in row])
            if search_lower in row_str:
                self.last_search_index = i
                self.tree.selection_set(str(i))
                self.tree.see(str(i))
                self.tree.focus(str(i))
                return 1 
        
        if start_idx > 0:
            self.last_search_index = -1
            if messagebox.askyesno("Recherche", "Fin du fichier atteinte. Recommencer au début ?", parent=self):
                return self.find_next_and_replace(search_text, replace_text, False)
        return 0

    def refresh_row(self, index):
        if str(index) in self.tree.get_children():
            row = self.data[index]
            display_cols = [c for c in self.headers if c in self.visible_columns]
            values = []
            for col in display_cols:
                try:
                    idx = self.headers.index(col)
                    val = row[idx] if idx < len(row) else ""
                    values.append(val)
                except ValueError:
                    values.append("")
            self.tree.item(str(index), values=values)

    def replace_all(self, search_text, replace_text):
        count = 0
        search_lower = search_text.lower()
        for r_idx, row in enumerate(self.data):
            changed = False
            for c_idx, cell in enumerate(row):
                cell_str = str(cell)
                if search_lower in cell_str.lower():
                    pattern = re.compile(re.escape(search_text), re.IGNORECASE)
                    new_val = pattern.sub(replace_text, cell_str)
                    if new_val != cell_str:
                        self.data[r_idx][c_idx] = new_val
                        changed = True
            if changed:
                count += 1
        if count > 0:
            self.refresh_tree()
        return count
    
    # =========================================================================
    #  TABLEWIDGET (COMPARAISON) - MENU CONTEXTUEL & INSERTION
    # =========================================================================
    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id: return

        if row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
            
        try:
            col_num = int(col_id.replace('#', '')) - 1
        except: return

        display_cols = [c for c in self.headers if c in self.visible_columns]
        if col_num < 0 or col_num >= len(display_cols): return
        col_name = display_cols[col_num]
        
        menu = tk.Menu(self, tearoff=0)
        
        # --- 1. COPIE VERTICALE (NOUVEAU) ---
        # Cette option copie uniquement les cellules de CETTE colonne pour les lignes sélectionnées
        menu.add_command(label=f"Copier colonne '{col_name}' (Sélection)", 
                         command=lambda: self.copy_block_to_clipboard(col_name, col_name))

        # --- 2. COPIE DE BLOC (HORIZONTAL) ---
        copy_submenu = tk.Menu(menu, tearoff=0)
        idx_start = display_cols.index(col_name)
        # On commence à +1 car la colonne elle-même est déjà gérée par l'option du dessus
        for i in range(idx_start + 1, min(idx_start + 40, len(display_cols))):
            end_c = display_cols[i]
            copy_submenu.add_command(label=f"Jusqu'à {end_c}", 
                                     command=lambda c=col_name, e=end_c: self.copy_block_to_clipboard(c, e))
        
        menu.add_cascade(label=f"Copier le bloc depuis '{col_name}'...", menu=copy_submenu)
        
        menu.add_separator()
        
        # --- 3. INSERTION ---
        insert_menu = tk.Menu(menu, tearoff=0)
        counts = [1, 2, 3, 4, 5, 10, 20, 50, 100]
        
        above = tk.Menu(insert_menu, tearoff=0)
        for n in counts:
            above.add_command(label=f"{n} ligne(s)", command=lambda c=n: self.insert_rows(row_id, c, 'above'))
        insert_menu.add_cascade(label="Insérer au-dessus", menu=above)
        
        below = tk.Menu(insert_menu, tearoff=0)
        for n in counts:
            below.add_command(label=f"{n} ligne(s)", command=lambda c=n: self.insert_rows(row_id, c, 'below'))
        insert_menu.add_cascade(label="Insérer en-dessous", menu=below)
        
        menu.add_cascade(label="Insérer des lignes...", menu=insert_menu)

        menu.add_separator()

        # --- 4. COLLER ---
        menu.add_command(label="Coller (Écraser)", 
                         command=lambda: self.paste_from_clipboard(row_id, col_name))
        
        menu.add_separator()
        menu.add_command(label=f"Propager '{col_name}'", command=lambda: self.apply_bulk_edit(row_id, col_name, "copy"))
        menu.add_command(label=f"Incrémenter '{col_name}'", command=lambda: self.apply_bulk_edit(row_id, col_name, "increment"))
        
        menu.add_separator()
        menu.add_command(label=f"Rechercher/Remplacer dans '{col_name}' (Sélection)", 
                         command=lambda: self.open_search_replace_popup(col_name))
        
        menu.tk_popup(event.x_root, event.y_root)

    def open_search_replace_popup(self, col_name):
        """Ouvre une pop-up de recherche/remplacement sur la sélection."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Attention", "Aucune ligne sélectionnée.")
            return

        # Création de la fenêtre
        top = tk.Toplevel(self)
        top.title(f"Remplacer dans : {col_name}")
        top.geometry("400x180")
        top.transient(self) # Reste au dessus
        top.resizable(False, False)
        
        # UI
        tk.Label(top, text="Rechercher :").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        entry_find = tk.Entry(top, width=30)
        entry_find.grid(row=0, column=1, padx=10, pady=10)
        entry_find.focus_set()

        tk.Label(top, text="Remplacer par :").grid(row=1, column=0, padx=10, pady=10, sticky='e')
        entry_replace = tk.Entry(top, width=30)
        entry_replace.grid(row=1, column=1, padx=10, pady=10)

        # Variables d'état pour le bouton "Suivant"
        col_idx = self.headers.index(col_name)
        # On convertit les IDs treeview en index entiers triés
        self.search_indices = sorted([int(item) for item in selected_items])
        self.current_search_pos = 0

        def do_replace_next():
            find_str = entry_find.get()
            repl_str = entry_replace.get()
            
            if not find_str: return

            # On cherche la prochaine occurrence à partir de la position actuelle
            start_pos = self.current_search_pos
            match_found = False
            
            for i in range(start_pos, len(self.search_indices)):
                r_idx = self.search_indices[i]
                # Vérification validité index
                if r_idx < len(self.data) and col_idx < len(self.data[r_idx]):
                    current_val = str(self.data[r_idx][col_idx])
                    
                    if find_str in current_val:
                        # Remplacement (Undo snapshot possible ici si besoin, mais lourd pour du pas à pas)
                        new_val = current_val.replace(find_str, repl_str)
                        self.data[r_idx][col_idx] = new_val
                        
                        # Update visuel
                        if str(r_idx) in self.tree.get_children():
                            vals = list(self.tree.item(str(r_idx), 'values'))
                            # On doit trouver l'index visuel
                            display_cols = [c for c in self.headers if c in self.visible_columns]
                            if col_name in display_cols:
                                v_idx = display_cols.index(col_name)
                                vals[v_idx] = new_val
                                self.tree.item(str(r_idx), values=vals)
                                self.tree.see(str(r_idx)) # Scroll vers l'élément
                                self.tree.selection_set(str(r_idx)) # Focus visuel
                        
                        self.current_search_pos = i + 1 # Prêt pour le suivant
                        match_found = True
                        break # On s'arrête à une modification
            
            if not match_found:
                messagebox.showinfo("Fin", "Aucune autre occurrence trouvée dans la sélection.", parent=top)
                self.current_search_pos = 0 # On boucle ou on arrête

        def do_replace_all():
            find_str = entry_find.get()
            repl_str = entry_replace.get()
            if not find_str: return

            count = 0
            # Pour TableWidget, on peut ajouter save_full_state_for_undo() ici si implémenté
            if hasattr(self, 'save_full_state_for_undo'):
                self.save_full_state_for_undo()

            for r_idx in self.search_indices:
                if r_idx < len(self.data) and col_idx < len(self.data[r_idx]):
                    current_val = str(self.data[r_idx][col_idx])
                    if find_str in current_val:
                        self.data[r_idx][col_idx] = current_val.replace(find_str, repl_str)
                        count += 1
            
            if count > 0:
                self.refresh_tree()
                messagebox.showinfo("Succès", f"{count} occurrences remplacées.", parent=top)
                top.destroy()
            else:
                messagebox.showinfo("Info", "Aucune occurrence trouvée.", parent=top)

        # Boutons
        btn_frame = tk.Frame(top)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)

        tk.Button(btn_frame, text="Exécuter / Suivant", command=do_replace_next, width=15).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Remplacer tout", command=do_replace_all, width=15).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Fermer", command=top.destroy, width=10).pack(side='left', padx=5)
    # =========================================================================
    #  GESTION UNDO & INSERTION POUR TABLEWIDGET
    # =========================================================================

    def save_full_state_for_undo(self):
        """Sauvegarde une copie complète pour le Ctrl+Z."""
        snapshot = [row[:] for row in self.data]
        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)

    def undo(self, event=None):
        """Restaure l'état précédent."""
        if not self.undo_stack:
            return

        last_action = self.undo_stack.pop()
        
        # Restauration brutale (Snapshot)
        self.data = last_action
        
        # Reset affichage
        self.filtered_indices = list(range(len(self.data)))
        self.refresh_tree()
        messagebox.showinfo("Undo", "Action annulée.", parent=self)

    def insert_rows(self, target_row_id, count, position='below'):
        """Insère des lignes vides."""
        try:
            self.save_full_state_for_undo() # Sauvegarde avant modif

            target_idx = int(target_row_id)
            insert_idx = target_idx if position == 'above' else target_idx + 1
            
            empty_row = [""] * len(self.headers)
            new_rows = [list(empty_row) for _ in range(count)]
            
            self.data[insert_idx:insert_idx] = new_rows
            
            self.filtered_indices = list(range(len(self.data)))
            self.refresh_tree()
            
            try:
                self.tree.see(str(insert_idx))
                self.tree.selection_set(str(insert_idx))
            except: pass
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'insérer : {e}", parent=self)

    def copy_block_to_clipboard(self, start_col, end_col):
        """Copie les lignes sélectionnées de start_col à end_col."""
        try:
            selected_items = self.tree.selection()
            if not selected_items: return

            display_cols = [c for c in self.headers if c in self.visible_columns]
            idx_s = display_cols.index(start_col)
            idx_e = display_cols.index(end_col)
            
            lines = []
            for item in selected_items:
                r_idx = int(item)
                row_vals = []
                # On boucle de la colonne de début à la colonne de fin
                for i in range(idx_s, idx_e + 1):
                    c_name = display_cols[i]
                    c_idx = self.headers.index(c_name)
                    # Sécurité si la ligne est plus courte que prévu
                    val = str(self.data[r_idx][c_idx]) if c_idx < len(self.data[r_idx]) else ""
                    row_vals.append(val)
                lines.append("\t".join(row_vals))
            
            # Envoi au presse-papier
            final_text = "\n".join(lines)
            self.clipboard_clear()
            self.clipboard_append(final_text)
            self.update() # IMPORTANT : Force la mise à jour immédiate
            
            messagebox.showinfo("Succès", f"Bloc copié !\n({len(selected_items)} lignes x {idx_e - idx_s + 1} colonnes)", parent=self)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur de copie : {e}", parent=self)

    def paste_from_clipboard(self, start_row_id, start_col_name):
        """Colle avec support Undo."""
        try:
            self.save_full_state_for_undo() # <--- AJOUT CRUCIAL

            content = self.clipboard_get()
            rows_to_paste = [line.split('\t') for line in content.splitlines()]
            if not rows_to_paste: return

            start_r = int(start_row_id)
            start_c = self.headers.index(start_col_name)

            for r_off, row_data in enumerate(rows_to_paste):
                curr_r = start_r + r_off
                if curr_r >= len(self.data): break
                for c_off, value in enumerate(row_data):
                    curr_c = start_c + c_off
                    if curr_c >= len(self.headers): break
                    val = value.strip()
                    if val.startswith('"') and val.endswith('"'): val = val[1:-1]
                    
                    while len(self.data[curr_r]) <= curr_c: self.data[curr_r].append("")
                    self.data[curr_r][curr_c] = val

            self.refresh_tree()
            messagebox.showinfo("Succès", "Données collées.", parent=self)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur collage : {e}", parent=self)

    def apply_bulk_edit(self, source_item_id, col_name, mode="copy"):
        """Applique la modification de masse sur le tableau principal."""
        try:
            # 1. Index réel de la colonne
            if col_name not in self.headers: return
            col_index = self.headers.index(col_name)
            
            # 2. Valeur source
            row_index = int(source_item_id)
            source_value = self.data[row_index][col_index] if col_index < len(self.data[row_index]) else ""
            
            # Préparation pour incrémentation
            start_num = 0
            is_number = False
            prefix = ""
            
            if mode == "increment":
                if str(source_value).isdigit():
                    start_num = int(source_value)
                    is_number = True
                else:
                    # Gestion des suffixes (ex: Capteur_1 -> Capteur_2)
                    match = re.search(r'(\d+)$', str(source_value))
                    if match:
                        start_num = int(match.group(1))
                        prefix = str(source_value)[:match.start()]
                        is_number = "suffix"
                    else:
                        messagebox.showwarning("Erreur", "Valeur non numérique, impossible d'incrémenter.", parent=self.root)
                        return

            # 3. Application à la sélection
            selected_items = self.tree.selection()
            
            for i, item_id in enumerate(selected_items):
                target_idx = int(item_id) # L'ID du treeview correspond à l'index dans self.data
                
                # Calcul de la nouvelle valeur
                new_val = source_value
                if mode == "increment":
                    if is_number == True:
                        new_val = str(start_num + i)
                    elif is_number == "suffix":
                        new_val = f"{prefix}{start_num + i}"

                # A. Mise à jour des données (Mémoire)
                # On s'assure que la ligne est assez longue
                while len(self.data[target_idx]) <= col_index:
                    self.data[target_idx].append("")
                
                self.data[target_idx][col_index] = str(new_val)
                
                # B. Mise à jour visuelle (Treeview)
                # On récupère les valeurs actuelles affichées pour ne changer que la cellule cible
                current_values = list(self.tree.item(item_id, 'values'))
                
                # On doit trouver l'index VISUEL (car certaines colonnes peuvent être masquées)
                display_cols = [c for c in self.headers if c in self.visible_columns]
                if col_name in display_cols:
                    visual_index = display_cols.index(col_name)
                    if visual_index < len(current_values):
                        current_values[visual_index] = str(new_val)
                        self.tree.item(item_id, values=current_values)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la modification : {e}", parent=self.root)


class ComparisonWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Comparaison de Fichiers")
        self.geometry("1400x800")
        self.lift()
        self.focus_force()
        
        # --- Outils de Recherche Globale ---
        tools_frame = tk.Frame(self, bg="#ecf0f1", pady=10)
        tools_frame.pack(fill='x')
        
        tk.Label(tools_frame, text="Rechercher :", bg="#ecf0f1").pack(side='left', padx=5)
        self.entry_search = tk.Entry(tools_frame, width=20)
        self.entry_search.pack(side='left', padx=5)
        
        tk.Label(tools_frame, text="Remplacer par :", bg="#ecf0f1").pack(side='left', padx=5)
        self.entry_replace = tk.Entry(tools_frame, width=20)
        self.entry_replace.pack(side='left', padx=5)
        
        tk.Label(tools_frame, text="Sur :", bg="#ecf0f1").pack(side='left', padx=10)
        self.target_var = tk.StringVar(value="both")
        tk.Radiobutton(tools_frame, text="Gauche", variable=self.target_var, value="left", bg="#ecf0f1").pack(side='left')
        tk.Radiobutton(tools_frame, text="Droite", variable=self.target_var, value="right", bg="#ecf0f1").pack(side='left')
        #tk.Radiobutton(tools_frame, text="Les deux", variable=self.target_var, value="both", bg="#ecf0f1").pack(side='left')
        
        tk.Button(tools_frame, text="Exécuter / Suivant", command=self.perform_next, bg="#e67e22", fg="white").pack(side='left', padx=10)
        tk.Button(tools_frame, text="Remplacer Tout", command=self.perform_replace_all, bg="#c0392b", fg="white").pack(side='left', padx=10)

        # --- Paned Window pour les Tableaux ---
        paned = tk.PanedWindow(self, orient='horizontal', sashwidth=5, bg="#bdc3c7")
        paned.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.left_panel = TableWidget(paned, title="Fichier Gauche", accent_color="#2980b9")
        self.right_panel = TableWidget(paned, title="Fichier Droite", accent_color="#8e44ad")
        
        paned.add(self.left_panel, width=700)
        paned.add(self.right_panel, width=700)

    def perform_next(self):
        s = self.entry_search.get()
        r = self.entry_replace.get()
        mode = self.target_var.get()
        
        # MODIFICATION : On active le remplacement uniquement si le champ Remplacer n'est pas vide
        do_replace_action = (r != "")

        found = False
        if mode in ["left", "both"]:
            # On passe do_replace_action au lieu de True
            res = self.left_panel.find_next_and_replace(s, r, do_replace=do_replace_action)
            if res: found = True
            
        if not found and mode in ["right", "both"]:
            # Idem ici
            res = self.right_panel.find_next_and_replace(s, r, do_replace=do_replace_action)
            if res: found = True

        if not found and self.left_panel.last_search_index == -1 and self.right_panel.last_search_index == -1:
            messagebox.showinfo("Info", "Aucune occurrence trouvée.", parent=self)

    def perform_replace_all(self):
        s = self.entry_search.get()
        r = self.entry_replace.get()
        mode = self.target_var.get()
        
        if not s: return
        
        # MODIFICATION : Si le champ remplacement est vide, on arrête tout de suite.
        if r == "":
            messagebox.showinfo("Info", "Le champ 'Remplacer par' est vide. Aucun remplacement effectué.", parent=self)
            return

        total = 0
        if mode in ["left", "both"]:
            total += self.left_panel.replace_all(s, r)
        if mode in ["right", "both"]:
            total += self.right_panel.replace_all(s, r)
        
        messagebox.showinfo("Résultat", f"{total} remplacements effectués au total.", parent=self)
    

class DatEditor:
    # ================= CONSTANTES DE COULEURS & STYLES =================
    COLORS = {
        "bg_dark": "#2c3e50",       # Sidebar background
        "bg_light": "#ecf0f1",      # Main content background
        "accent": "#3498db",        # Primary button color (Active Module)
        "accent_hover": "#2980b9",  # Primary button hover
        "text_white": "#ffffff",
        "text_dark": "#2c3e50",
        "row_alt": "#f7f9f9",       # Zebra stripe row
        "white": "#ffffff",
        "success": "#27ae60",       # Green buttons
        "warning": "#f39c12",       # Orange/Yellow
        "danger": "#c0392b"         # Red
    }

    COMM_DEFAULT_HEADERS = [
        "Type", "Version", "Réseau", "Nom", "Equipement", "Type de trame",
        "Caractéristiques", "Quantité", "Lecture/Ecriture",
        "Adresse de début ou hh NETWORK", "Type de données ou mm NETWORK",
        "Période de scrutation (1 ou 0) ou ss NETWORK", "hh ou fff NETWORK", "mm ou actif au démarrage NETWORK (1 ou 0)", "ss", "fff",
        "-", "Numéro de DB", "Adresse IP ou 0",
        "Si EQT -> = 0", "--", "0 si NETWORK",
        "0 ou Descritpion si NETWORK", "Description ou 0 si NETWORK",
        "0 si NETWORK ou EQT", "3 si NETWORK", "1 si NETWORK", "Nom du protocole"
    ]

    EVENT_DEFAULT_HEADERS = [
        "Mode", "Nom", "Description", "00", "0", "Nom de liste serveurs", "Vide",
        "Variable scrutée", "Activation bit (0 = 1>0 ou 1 = 0>1 ou 2 = expression)", "1",
        "Variable bit activation", "Expression (si expression)", "Programme",
        "Branche", "Fonction", "Argument", "=1"
    ]
    
    EXPRV_DEFAULT_HEADERS = [
        "Mode", "Nom", "Description", "00", "0", "Nom de liste serveurs", "Vide",
        "1", "Variable activation", "Variable", "Branche", "Expression", "=1"
    ]
    
    CYCLIC_DEFAULT_HEADERS = [
        "Mode", "Nom", "Description", "00", "0", "Nom de liste serveurs", "Vide",
        "Nombre de secondes de cycle",
        "1 si bit d'activation 0 sinon",
        "Variable d'activation",
        "Programme", "Branche", "Fonction", "Argument",
        "=0", "=1"
    ]
    VARTREAT_DEFAULT_HEADERS = [
        "TREATMENT", "GROUPALARM", "Nom", "0", "Nom de liste serveurs", "Prise en compte de la population appliquée (0 ou vide)",
        "Description", "Filtre de branche (1 ou 0)", "Filtre de branche (branche)", "Vide 1", "Niveau d'alarme min", "Niveau d'alarme max", "Expression",
        "Variable Priorité d'alarme présente acquittée la plus haute", "Variable Priorité d'alarme présente non acquittée la plus haute",
        "Vide 2", "Vide 3", "Vide 4", "Nom de la branche", "Variable Nombre d'alarmes présentes non acquittées",
        "Variable Nombre d'alarmes présentes acquittées", "Variable Nombre d'alarmes présentes (acquittées ou non)",
        "Variable Nombre d'alarmes au repos non acquittées", "Variable Nombre d'alarmes au repos",
        "Variable Nombre d'alarmes invalides", "Variable Nombre d'alarmes masquées", "Variable Nombre d'alarmes masquées par utilisateur",
        "Variable Nombre d'alarmes masquées par programme",
        "Variable Nombre d'alarmes masquées par dépendance sur une autre variable",
        "Variable Nombre d'alarmes masquées par expression", "Variable Nombre d'alarmes présentes et en mode prise en compte",
        "Variable Nombre d'alarmes au repos et en mode prise en compte", "Nombre d'alarmes inhibées"
    ]
    
    VAREXP_TEMPLATES = {
        "CMD": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "Log0_1": "0", "Log1_0": "0", "BitCommandLevel": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "BIT": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "Log0_1": "0", "Log1_0": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmAcknowledgmentLevel": "-1",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "ACM": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "Log0_1": "0", "Log1_0": "0", "BitCommandLevel": "0",
            "AlarmLevel": "0", "AlarmActiveAt1": "1",
            "AlarmTemporization": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "ALA": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "Log0_1": "0", "Log1_0": "0",
            "AlarmLevel": "0", "AlarmActiveAt1": "1",
            "AlarmTemporization": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "REG": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "DeadbandValue": "0", "MinimumValue": "0",
            "ScaledValue": "0", "DeviceMinimumValue": "0",
            "MaximumValue": "65535", "DeviceMaximumValue": "65535",
            "DeadbandType": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "CTV": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "DeadbandValue": "0", "MinimumValue": "0",
            "ScaledValue": "0", "DeviceMinimumValue": "0",
            "MaximumValue": "65535", "DeviceMaximumValue": "65535",
            "ControlMinimumValue": "0", "RegisterCommandLevel": "0",
            "ControlMaximumValue": "65535",
            "DeadbandType": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "TXT": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "Textsize": "132", "TextCommandLevel": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "CXT": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "Textsize": "132", "TextCommandLevel": "0",
            "ExtBinary": "0", "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        },
        "CHR": {
            "Source": "I", "Broadcast": "0", "StationOrAssociationNumber": "0",
            "UseExtendedAttributes": "1", "PermanentScan": "2",
            "MinimumValue": "0", "ScaledValue": "0", "DeviceMinimumValue": "0",
            "MaximumValue": "65535", "DeviceMaximumValue": "65535", "Chrono_Period": "100",
            "Chrono_Type": "1", "Chrono_EnableBitName": "", "Chrono_EnableBitTransition": "1", 
            "Chrono_ResetBitName": "", "Chrono_ResetBitTransition": "1", "DeadbandType": "0",
            "Recorder": "0", "MessageAlarm": "0",
            "BrowsingLevel": "0", "WithInitialValue": "0", "InitialValue": "0",
            "AlarmMaskLevel": "-2", "AlarmMaintenanceLevel": "-2"
        }
    }
    VAREXP_DESCRIPTIONS = {
        "CMD": "Etat commandable",
        "BIT": "Etat",
        "ACM": "Alarme commandable",
        "ALA": "Alarme",
        "REG": "Registre analogique",
        "CTV": "Registre analogique commandable",
        "TXT": "Texte",
        "CXT": "Texte commandable"
    }


    def __init__(self, root):
        self.root = root
        self.root.title("Éditeur .DAT - Modernisé (Optimisé)")
        self.root.geometry("1920x1080")
        self.root.configure(bg=self.COLORS["bg_light"])
        
        # Configuration des polices
        self.default_font = tkfont.Font(family="Segoe UI", size=10)
        self.header_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.title_font = tkfont.Font(family="Segoe UI", size=12, weight="bold")

        # === VARIABLES DE PAGINATION (NOUVEAU) ===
        self.view_start = 0      # Index de départ
        self.view_limit = 1000    # Nombre de lignes affichées (taille de la fenêtre)
        self.view_step = 250     # Décalage lors du clic (overlap)
        
        # Configuration du style global
        self._configure_styles()
    
        # Données
        self.data = []
        self.headers = []
        self.filtered_indices = []
        self.visible_columns = []
        self.first_line = None
        self.selected_folder = None
        self.modified = False
    
        # Pour copier/coller
        self.clipboard_rows = []
    
        # Undo stack
        self.undo_stack = []
        # self.copy_undo_stack_size = None # SUPPRIMÉ POUR CORRECTION BUG UNDO
        
        # ---- Frame principale ----
        main_frame = tk.Frame(root, bg=self.COLORS["bg_light"])
        main_frame.pack(fill='both', expand=True)
    
        # ---- Barre verticale (Sidebar - Navigation) ----
        nav_frame = tk.Frame(main_frame, bg=self.COLORS["bg_dark"], width=240)
        nav_frame.pack(side='left', fill='y')
        nav_frame.pack_propagate(False) # Force la largeur
        
        # Titre dans la sidebar
        tk.Label(nav_frame, text="Editeur .DAT", bg=self.COLORS["bg_dark"], fg=self.COLORS["text_white"], 
                 font=("Segoe UI", 16, "bold"), pady=20).pack(fill='x')

        # Conteneur pour les boutons de navigation (Sidebar)
        nav_content = tk.Frame(nav_frame, bg=self.COLORS["bg_dark"])
        nav_content.pack(fill='both', expand=True, padx=10)

        # ---- Contenu Principal (Droite) ----
        content_frame = tk.Frame(main_frame, bg=self.COLORS["bg_light"])
        content_frame.pack(side='left', fill='both', expand=True, padx=20, pady=20)
    
        # ---- Boutons Sidebar (Utilisation d'une méthode helper pour le style) ----
        self.buttons = {}
        
        # Groupe Fichiers
        self._add_nav_label(nav_content, "Fichiers")
        self.buttons['folder'] = self._create_nav_button(nav_content, "Choisir dossier", self.select_folder)
        self.buttons['open_any'] = self._create_nav_button(nav_content, "Ouvrir autre .DAT", self.open_any_dat_file)
        self.buttons['save'] = self._create_nav_button(nav_content, "Enregistrer sous...", self.save_file, bg_color=self.COLORS["success"])
        
        self._add_nav_separator(nav_content)
        
        # Groupe Modules
        self._add_nav_label(nav_content, "Modules")
        self.buttons['varexp'] = self._create_nav_button(nav_content, "Varexp (Variables)", self.load_varexp)
        self.buttons['comm'] = self._create_nav_button(nav_content, "Comm (Réseau)", self.load_comm)
        self.buttons['event'] = self._create_nav_button(nav_content, "Event", self.load_event)
        self.buttons['exprv'] = self._create_nav_button(nav_content, "Exprv", self.load_exprv)
        self.buttons['cyclic'] = self._create_nav_button(nav_content, "Cyclic", self.load_cyclic)
        self.buttons['vartreat'] = self._create_nav_button(nav_content, "Vartreat", self.load_vartreat)
        
        self._add_nav_separator(nav_content)

        # Groupe Création
        self._add_nav_label(nav_content, "Création")
        self.buttons['create_var'] = self._create_nav_button(nav_content, "Créer Variable", self.open_create_variable, state="disabled")
        self.buttons['create_event'] = self._create_nav_button(nav_content, "Créer Event", lambda: self.open_create_generic('event'), state="disabled")
        self.buttons['create_exprv'] = self._create_nav_button(nav_content, "Créer Expression", lambda: self.open_create_generic('exprv'), state="disabled")
        self.buttons['create_cyclic'] = self._create_nav_button(nav_content, "Créer Cyclic", lambda: self.open_create_generic('cyclic'), state="disabled")
        self.buttons['create_vartreat'] = self._create_nav_button(nav_content, "Créer Synthèse", lambda: self.open_create_generic('vartreat'), state="disabled")
        
        # ================= ZONE SUPÉRIEURE : Filtres (Gauche) + Outils (Droite) =================
        top_container = tk.Frame(content_frame, bg=self.COLORS["bg_light"])
        top_container.pack(fill='x', pady=(0, 15))

        # ---- 1. Zone de Filtrage (Gauche - Etirée) ----
        filter_card = tk.Frame(top_container, bg="white", highlightthickness=1, highlightbackground="#dcdcdc")
        filter_card.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        tk.Label(filter_card, text="Filtres", font=self.title_font, bg="white", fg=self.COLORS["accent"]).pack(anchor='w', padx=15, pady=5)
        
        frame_filter = tk.Frame(filter_card, bg="white")
        frame_filter.pack(fill='x', padx=15, pady=(0, 10))
        
        # Configuration de la grille pour que les filtres s'étirent
        frame_filter.columnconfigure(0, weight=1)
        frame_filter.columnconfigure(1, weight=1)
        frame_filter.columnconfigure(2, weight=1)
        frame_filter.columnconfigure(3, weight=0) # Logic boutons
        frame_filter.columnconfigure(4, weight=0) # Action boutons
    
        # Helper pour les widgets de filtre (GRID)
        def create_filter_group(parent, text, col_idx):
            f = tk.Frame(parent, bg="white")
            f.grid(row=0, column=col_idx, sticky="ew", padx=5)
            
            tk.Label(f, text=text, bg="white", font=("Segoe UI", 8, "bold"), fg="#7f8c8d").pack(anchor='w')
            
            cb = ttk.Combobox(f, state="readonly")
            cb.pack(fill='x', pady=1)
            
            entry = ttk.Entry(f)
            entry.pack(fill='x', pady=1)
            
            # Ajout du bind Entrée pour appliquer le filtre rapidement
            entry.bind("<Return>", lambda e: self.apply_filter())
            
            return cb, entry

        self.column_filter1, self.filter_entry1 = create_filter_group(frame_filter, "Colonne 1", 0)
        self.column_filter2, self.filter_entry2 = create_filter_group(frame_filter, "Colonne 2", 1)
        self.column_filter3, self.filter_entry3 = create_filter_group(frame_filter, "Colonne 3", 2)
    
        # Logic buttons styled
        logic_frame = tk.Frame(frame_filter, bg="white")
        logic_frame.grid(row=0, column=3, padx=10, sticky="ns")
        
        self.logic_mode = tk.StringVar(value="ET")
        ttk.Radiobutton(logic_frame, text="ET", variable=self.logic_mode, value="ET").pack(anchor='w')
        ttk.Radiobutton(logic_frame, text="OU", variable=self.logic_mode, value="OU").pack(anchor='w')
        
        # Action buttons for filter
        btn_filter_frame = tk.Frame(frame_filter, bg="white")
        btn_filter_frame.grid(row=0, column=4, padx=5, sticky="ns")
        
        tk.Button(btn_filter_frame, text="Appliquer", bg=self.COLORS["accent"], fg="white", 
                  relief="flat", font=("Segoe UI", 8, "bold"), padx=10, pady=2, width=10,
                  command=self.apply_filter).pack(pady=1)
        tk.Button(btn_filter_frame, text="Reset", bg="#95a5a6", fg="white", 
                  relief="flat", font=("Segoe UI", 8), padx=10, pady=2, width=10,
                  command=self.reset_filters).pack(pady=1)

        # ---- 2. Zone Outils (Droite) ----
        tools_card = tk.Frame(top_container, bg="white", highlightthickness=1, highlightbackground="#dcdcdc")
        tools_card.pack(side='right', fill='y', padx=(0, 0)) 

        tk.Label(tools_card, text="Outils", font=self.title_font, bg="white", fg="#7f8c8d").pack(anchor='w', padx=15, pady=5)
        
        tools_inner = tk.Frame(tools_card, bg="white")
        tools_inner.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        def create_tool_button(parent, text, cmd, color=self.COLORS["bg_dark"]):
            btn = tk.Button(parent, text=text, command=cmd, bg=color, fg="white", 
                            relief="flat", font=("Segoe UI", 9), padx=10, pady=5, width=12)
            btn.pack(side='left', padx=2)
            return btn

        self.buttons['columns'] = create_tool_button(tools_inner, "Colonnes", self.select_columns)
        self.buttons['search'] = create_tool_button(tools_inner, "Rechercher", self.open_search_replace)
        self.buttons['top'] = create_tool_button(tools_inner, "▲ Haut", self.scroll_top, color="#95a5a6")
        self.buttons['bottom'] = create_tool_button(tools_inner, "▼ Bas", self.scroll_bottom, color="#95a5a6")
        self.buttons['compare'] = create_tool_button(tools_inner, "Comparaison", self.open_comparison, color="#8e44ad")

    
        # ---- Table + Scrollbars ----
        table_container = tk.Frame(content_frame, bg="white", highlightthickness=1, highlightbackground="#dcdcdc")
        table_container.pack(fill='both', expand=True)
        
        frame_table_parent = tk.Frame(table_container, bg="white")
        frame_table_parent.pack(fill='both', expand=True, padx=1, pady=1)
        
        frame_table_parent.columnconfigure(0, weight=1)
        frame_table_parent.columnconfigure(1, weight=0)
        frame_table_parent.rowconfigure(0, weight=1)
    
        # Treeview
        self.frame_table = tk.Frame(frame_table_parent)
        self.frame_table.grid(row=0, column=0, sticky='nsew')
    
        self.vsb = ttk.Scrollbar(self.frame_table, orient='vertical')
        self.hsb = ttk.Scrollbar(self.frame_table, orient='horizontal')
    
        self.tree = ttk.Treeview(
            self.frame_table,
            show='headings',
            selectmode='extended',
            yscrollcommand=self.vsb.set,
            xscrollcommand=self.hsb.set,
            style="Custom.Treeview"
        )
        self.tree.grid(row=0, column=0, sticky='nsew')
    
        self.vsb.config(command=self.tree.yview)
        self.vsb.grid(row=0, column=1, sticky='ns')
        self.hsb.config(command=self.tree.xview)
        self.hsb.grid(row=1, column=0, sticky='ew')
    
        self.frame_table.columnconfigure(0, weight=1)
        self.frame_table.rowconfigure(0, weight=1)
    
        # Configurer les tags pour les couleurs alternées
        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background=self.COLORS["row_alt"])

        # Bande droite fixe (contexte visuel)
        self.right_band = tk.Frame(frame_table_parent, width=20, bg="#ecf0f1")
        self.right_band.grid(row=0, column=1, sticky='ns')
    
        # ---- Binds clavier et double-clic ----
        self.root.bind("<Control-c>", self.copy_rows)
        self.root.bind("<Control-v>", self.paste_rows)
        self.root.bind("<Control-z>", self.undo)
        self.root.bind("<Delete>", self.delete_selected_rows)
        self.tree.bind('<Double-1>', self.edit_cell)
        self.tree.bind('<Button-3>', self.show_context_menu) # Windows / Linux
        self.tree.bind('<Button-2>', self.show_context_menu) # MacOS
    
        # ================= BARRE DE STATUT =================
        self.status_var = tk.StringVar()
        self.status_var.set("Prêt.")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            bg=self.COLORS["accent"],
            fg="white",
            padx=15,
            pady=5,
            font=("Segoe UI", 9)
        )
        status_bar.pack(side="bottom", fill="x")
    
        # Mises à jour automatiques
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.update_status_bar())
        self.tree.bind("<Motion>", lambda e: self.update_status_bar())
        
        self.editing_entry = None  
        self.editing_item = None   
        self.editing_header_idx = None
        self.sort_state = {}
        self.saved_advanced_params = {} 
        
        # Fonction pour repositionner l'Entry si nécessaire
        def reposition_entry(event=None):
            if self.editing_entry and self.editing_item is not None:
                tree_index, col_index = self.editing_item
                item_iid = str(tree_index)
                column_id = f"#{col_index+1}"
                bbox = self.tree.bbox(item_iid, column_id)
                if not bbox:
                    self.editing_entry.place_forget()
                    return
                x, y, w, h = bbox
                # Ajustement pour le style modernisé
                self.editing_entry.place(
                    x=self.tree.winfo_rootx() - self.root.winfo_rootx() + x,
                    y=self.tree.winfo_rooty() - self.root.winfo_rooty() + y,
                    width=w, height=h
                )
        
        self.tree.bind("<MouseWheel>", reposition_entry)
        self.tree.bind("<Button-4>", reposition_entry)
        self.tree.bind("<Button-5>", reposition_entry)
        self.tree.bind("<Shift-MouseWheel>", reposition_entry)
        self.tree.bind("<Configure>", reposition_entry)
        self.tree.bind("<Motion>", reposition_entry)
        self.vsb.bind("<B1-Motion>", reposition_entry)
        self.hsb.bind("<B1-Motion>", reposition_entry)
        
        self.last_tagname = 0  
        self.cell_templates = {} 
        
        for idx, row in enumerate(self.data):
            iid = str(idx)
            self.cell_templates[iid] = {col: row[self.headers.index(col)] for col in self.visible_columns if col in self.headers}

    def open_comparison(self):
        # 1. Sécurité : On vérifie si la variable existe, sinon on la crée
        if not hasattr(self, 'comparison_window'):
            self.comparison_window = None

        # 2. Si la fenêtre semble ouverte, on essaie de la mettre au premier plan
        if self.comparison_window is not None:
            try:
                if self.comparison_window.winfo_exists():
                    self.comparison_window.lift()
                    self.comparison_window.focus_force()
                    return
                else:
                    # Elle n'existe plus (fermée sauvagement), on réinitialise
                    self.comparison_window = None
            except Exception:
                self.comparison_window = None

        # 3. Création de la nouvelle fenêtre
        try:
            self.comparison_window = ComparisonWindow(self.root)
            
            # Gestion propre de la fermeture (croix rouge)
            def on_close():
                try:
                    self.comparison_window.destroy()
                except:
                    pass
                self.comparison_window = None 

            self.comparison_window.protocol("WM_DELETE_WINDOW", on_close)
            
        except NameError:
            messagebox.showerror("Erreur", "La classe 'ComparisonWindow' est introuvable.\nVérifiez qu'elle est bien collée AVANT la classe DatEditor.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir la fenêtre : {e}")
        
    # =========================================================================
    #  DATEDITOR (PRINCIPAL) - MENU, INSERTION & UNDO
    # =========================================================================
    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id: return

        if row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
        
        try:
            col_num = int(col_id.replace('#', '')) - 1
        except: return

        display_cols = [c for c in self.headers if c in self.visible_columns]
        if col_num < 0 or col_num >= len(display_cols): return
        col_name = display_cols[col_num]

        menu = tk.Menu(self.root, tearoff=0)

        # --- 1. COPIE VERTICALE (NOUVEAU) ---
        # Astuce : On appelle copy_block avec start_col == end_col
        menu.add_command(label=f"Copier colonne '{col_name}' (Sélection)", 
                         command=lambda: self.copy_block_to_clipboard(col_name, col_name))

        # --- 2. COPIE DE BLOC ---
        copy_submenu = tk.Menu(menu, tearoff=0)
        idx_start = display_cols.index(col_name)
        for i in range(idx_start + 1, min(idx_start + 40, len(display_cols))):
            end_c = display_cols[i]
            copy_submenu.add_command(label=f"Jusqu'à {end_c}", 
                                     command=lambda c=col_name, e=end_c: self.copy_block_to_clipboard(c, e))
        menu.add_cascade(label=f"Copier le bloc depuis '{col_name}'...", menu=copy_submenu)
        
        menu.add_separator()

        # --- 3. INSERTION ---
        insert_menu = tk.Menu(menu, tearoff=0)
        counts = [1, 2, 3, 4, 5, 10, 20, 50, 100]
        
        above_menu = tk.Menu(insert_menu, tearoff=0)
        for n in counts:
            above_menu.add_command(label=f"{n} ligne(s)", command=lambda c=n: self.insert_rows(row_id, c, 'above'))
        insert_menu.add_cascade(label="Insérer au-dessus", menu=above_menu)
        
        below_menu = tk.Menu(insert_menu, tearoff=0)
        for n in counts:
            below_menu.add_command(label=f"{n} ligne(s)", command=lambda c=n: self.insert_rows(row_id, c, 'below'))
        insert_menu.add_cascade(label="Insérer en-dessous", menu=below_menu)
        
        menu.add_cascade(label="Insérer des lignes...", menu=insert_menu)
        menu.add_separator()

        # --- 4. COLLER ---
        menu.add_command(label="Coller (Écraser)", 
                         command=lambda: self.paste_from_clipboard(row_id, col_name))
        
        menu.add_separator()
        menu.add_command(label=f"Propager '{col_name}'", command=lambda: self.apply_bulk_edit(row_id, col_name, "copy"))
        menu.add_command(label=f"Incrémenter '{col_name}'", command=lambda: self.apply_bulk_edit(row_id, col_name, "increment"))
        
        menu.add_separator()
        menu.add_command(label=f"Rechercher/Remplacer dans '{col_name}'", 
                         command=lambda: self.open_search_replace_popup(col_name))
        
        menu.tk_popup(event.x_root, event.y_root)

    def open_search_replace_popup(self, col_name):
        """Pop-up Rechercher/Remplacer ciblée sur la colonne et les lignes sélectionnées."""
        selected_items = self.tree.selection()
        if not selected_items: return

        top = tk.Toplevel(self.root)
        top.title(f"Remplacer dans la colonne : {col_name}")
        top.geometry("420x180")
        top.transient(self.root)
        top.resizable(False, False)
        
        # Layout
        tk.Label(top, text="Rechercher :").grid(row=0, column=0, padx=10, pady=15, sticky='e')
        entry_find = tk.Entry(top, width=35)
        entry_find.grid(row=0, column=1, padx=10, pady=15)
        entry_find.focus_set()

        tk.Label(top, text="Remplacer par :").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        entry_replace = tk.Entry(top, width=35)
        entry_replace.grid(row=1, column=1, padx=10, pady=5)

        # Données
        col_idx = self.headers.index(col_name)
        self.search_indices = sorted([int(item) for item in selected_items])
        self.current_search_pos = 0

        def do_replace_next():
            """Remplace la PROCHAINE occurrence trouvée."""
            find_str = entry_find.get()
            repl_str = entry_replace.get()
            if not find_str: return

            match_found = False
            
            # On parcourt la sélection à partir du dernier point
            for i in range(self.current_search_pos, len(self.search_indices)):
                r_idx = self.search_indices[i]
                if r_idx < len(self.data):
                    current_val = str(self.data[r_idx][col_idx])
                    
                    if find_str in current_val:
                        # Petite sauvegarde undo unitaire pour le "pas à pas"
                        self.undo_stack.append([(r_idx, col_idx, current_val)]) 
                        
                        # Remplacement
                        new_val = current_val.replace(find_str, repl_str, 1) # Remplace 1ère occurrence ou toutes ? Généralement toutes dans la cellule
                        new_val = current_val.replace(find_str, repl_str)
                        self.data[r_idx][col_idx] = new_val
                        
                        # Update Visuel
                        if str(r_idx) in self.tree.get_children():
                            vals = list(self.tree.item(str(r_idx), 'values'))
                            # Trouver index visuel
                            display_cols = [c for c in self.headers if c in self.visible_columns]
                            if col_name in display_cols:
                                v_idx = display_cols.index(col_name)
                                vals[v_idx] = new_val
                                self.tree.item(str(r_idx), values=vals)
                                self.tree.see(str(r_idx))
                                self.tree.selection_set(str(r_idx))
                        
                        self.modified = True
                        self.current_search_pos = i + 1
                        match_found = True
                        break # On stop pour attendre le prochain clic
            
            if not match_found:
                messagebox.showinfo("Fin", "Terminé pour la sélection.", parent=top)
                self.current_search_pos = 0

        def do_replace_all():
            """Remplace TOUT dans la sélection."""
            find_str = entry_find.get()
            repl_str = entry_replace.get()
            if not find_str: return

            count = 0
            # SAUVEGARDE UNDO MASSIVE
            self.save_full_state_for_undo()

            for r_idx in self.search_indices:
                if r_idx < len(self.data):
                    current_val = str(self.data[r_idx][col_idx])
                    if find_str in current_val:
                        self.data[r_idx][col_idx] = current_val.replace(find_str, repl_str)
                        count += 1
            
            if count > 0:
                self.modified = True
                self.refresh_tree() # Refresh global pour être sûr
                messagebox.showinfo("Succès", f"{count} remplacements effectués.", parent=top)
                top.destroy()
            else:
                messagebox.showinfo("Info", "Aucune correspondance trouvée.", parent=top)

        # Boutons
        btn_frame = tk.Frame(top)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)

        # Style bouton
        tk.Button(btn_frame, text="Exécuter / Suivant", command=do_replace_next, bg="#3498db", fg="white", relief="flat").pack(side='left', padx=5, ipady=3)
        tk.Button(btn_frame, text="Remplacer tout", command=do_replace_all, bg="#e74c3c", fg="white", relief="flat").pack(side='left', padx=5, ipady=3)
        tk.Button(btn_frame, text="Fermer", command=top.destroy, relief="flat").pack(side='left', padx=5, ipady=3)
        
    def insert_rows(self, target_row_id, count, position='below'):
        """Insère des lignes vides avec support Undo."""
        try:
            # 1. SAUVEGARDE UNDO AVANT MODIF
            # On utilise votre méthode existante ou on crée une snapshot manuelle
            self.save_full_state_for_undo() # Voir méthode ci-dessous si elle n'existe pas

            target_idx = int(target_row_id)
            insert_idx = target_idx if position == 'above' else target_idx + 1
            
            # 2. Création
            empty_row = [""] * len(self.headers)
            new_rows = [list(empty_row) for _ in range(count)]
            
            # 3. Insertion
            self.data[insert_idx:insert_idx] = new_rows
            self.modified = True
            
            # 4. Refresh complet nécessaire car les IDs changent
            self.filtered_indices = list(range(len(self.data)))
            self.refresh_tree()
            
            # 5. Scroll
            try:
                self.tree.see(str(insert_idx))
                self.tree.selection_set(str(insert_idx))
            except: pass
            
            self.status_var.set(f"{count} lignes insérées.")

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur insertion : {e}")

    def copy_block_to_clipboard(self, start_col, end_col):
        """Copie le bloc vers le presse-papier système."""
        try:
            selected_items = self.tree.selection()
            if not selected_items: return

            display_cols = [c for c in self.headers if c in self.visible_columns]
            idx_s = display_cols.index(start_col)
            idx_e = display_cols.index(end_col)
            
            lines = []
            for item in selected_items:
                r_idx = int(item)
                row_vals = []
                for i in range(idx_s, idx_e + 1):
                    c_name = display_cols[i]
                    c_idx = self.headers.index(c_name)
                    val = str(self.data[r_idx][c_idx]) if c_idx < len(self.data[r_idx]) else ""
                    row_vals.append(val)
                lines.append("\t".join(row_vals))
            
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(lines))
            self.root.update() # CRUCIAL : Force la validation du presse-papier
            
            self.status_var.set(f"Bloc copié : {len(lines)} lignes x {idx_e - idx_s + 1} colonnes")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Echec copie : {e}")

    def paste_from_clipboard(self, start_row_id, start_col_name):
        """Collage avec Undo."""
        try:
            # 1. SAUVEGARDE UNDO
            self.save_full_state_for_undo()

            content = self.root.clipboard_get()
            rows_to_paste = [line.split('\t') for line in content.splitlines()]
            if not rows_to_paste: return

            start_r = int(start_row_id)
            start_c = self.headers.index(start_col_name)
            display_cols = [c for c in self.headers if c in self.visible_columns]

            for r_off, row_data in enumerate(rows_to_paste):
                curr_r = start_r + r_off
                if curr_r >= len(self.data): break
                for c_off, value in enumerate(row_data):
                    curr_c = start_c + c_off
                    if curr_c >= len(self.headers): break
                    val = value.strip()
                    if val.startswith('"') and val.endswith('"'): val = val[1:-1]
                    
                    while len(self.data[curr_r]) <= curr_c: self.data[curr_r].append("")
                    self.data[curr_r][curr_c] = val
                    
                    col_real_name = self.headers[curr_c]
                    if col_real_name in display_cols:
                        v_idx = display_cols.index(col_real_name)
                        if self.tree.exists(str(curr_r)):
                            vals = list(self.tree.item(str(curr_r), 'values'))
                            if v_idx < len(vals):
                                vals[v_idx] = val
                                self.tree.item(str(curr_r), values=vals)
            
            self.modified = True
            self.status_var.set("Collage effectué (Undo possible).")
        except Exception as e:
            messagebox.showerror("Erreur", "Presse-papier invalide.")
            
    def save_full_state_for_undo(self):
        """
        À appeler AVANT une grosse modification (Coller bloc, Insérer lignes).
        Sauvegarde une copie complète des données.
        """
        # On fait une copie profonde (Deep Copy) des données
        snapshot = [row[:] for row in self.data]
        self.undo_stack.append(snapshot)
        
        # Limite de sécurité (ex: 20 derniers états) pour ne pas saturer la RAM
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)
            
    def apply_bulk_edit(self, source_item_id, col_name, mode="copy"):
        """Applique la modification de masse sur le tableau principal."""
        try:
            # 1. Index réel de la colonne
            if col_name not in self.headers: return
            col_index = self.headers.index(col_name)
            
            # 2. Valeur source
            row_index = int(source_item_id)
            source_value = self.data[row_index][col_index] if col_index < len(self.data[row_index]) else ""
            
            # Préparation pour incrémentation
            start_num = 0
            is_number = False
            prefix = ""
            
            if mode == "increment":
                if str(source_value).isdigit():
                    start_num = int(source_value)
                    is_number = True
                else:
                    # Gestion des suffixes (ex: Capteur_1 -> Capteur_2)
                    match = re.search(r'(\d+)$', str(source_value))
                    if match:
                        start_num = int(match.group(1))
                        prefix = str(source_value)[:match.start()]
                        is_number = "suffix"
                    else:
                        messagebox.showwarning("Erreur", "Valeur non numérique, impossible d'incrémenter.", parent=self.root)
                        return

            # 3. Application à la sélection
            selected_items = self.tree.selection()
            
            for i, item_id in enumerate(selected_items):
                target_idx = int(item_id) # L'ID du treeview correspond à l'index dans self.data
                
                # Calcul de la nouvelle valeur
                new_val = source_value
                if mode == "increment":
                    if is_number == True:
                        new_val = str(start_num + i)
                    elif is_number == "suffix":
                        new_val = f"{prefix}{start_num + i}"

                # A. Mise à jour des données (Mémoire)
                # On s'assure que la ligne est assez longue
                while len(self.data[target_idx]) <= col_index:
                    self.data[target_idx].append("")
                
                self.data[target_idx][col_index] = str(new_val)
                
                # B. Mise à jour visuelle (Treeview)
                # On récupère les valeurs actuelles affichées pour ne changer que la cellule cible
                current_values = list(self.tree.item(item_id, 'values'))
                
                # On doit trouver l'index VISUEL (car certaines colonnes peuvent être masquées)
                display_cols = [c for c in self.headers if c in self.visible_columns]
                if col_name in display_cols:
                    visual_index = display_cols.index(col_name)
                    if visual_index < len(current_values):
                        current_values[visual_index] = str(new_val)
                        self.tree.item(item_id, values=current_values)

        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la modification : {e}", parent=self.root)
            
    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use('clam') # 'clam' allows better color customization than 'vista' or 'winnative'
        except:
            pass

        # Treeview Styles
        style.configure("Custom.Treeview", 
                        background="white",
                        foreground="black", 
                        rowheight=30, 
                        fieldbackground="white",
                        font=("Segoe UI", 10),
                        borderwidth=0)
        
        style.configure("Custom.Treeview.Heading", 
                        font=("Segoe UI", 10, "bold"), 
                        background="#dfe6e9", 
                        foreground="#2c3e50",
                        relief="flat")
        
        style.map("Custom.Treeview", 
                  background=[('selected', self.COLORS['accent'])], 
                  foreground=[('selected', 'white')])

        style.map("Custom.Treeview.Heading", 
                  background=[('active', '#b2bec3')])
        
        # Combobox
        style.configure("TCombobox", padding=5)
        
        # Scrollbars
        style.configure("Vertical.TScrollbar", background="#bdc3c7", troughcolor="#ecf0f1", borderwidth=0)
        style.configure("Horizontal.TScrollbar", background="#bdc3c7", troughcolor="#ecf0f1", borderwidth=0)

    def _add_nav_label(self, parent, text):
        tk.Label(parent, text=text.upper(), bg=self.COLORS["bg_dark"], fg="#95a5a6", 
                 font=("Segoe UI", 8, "bold"), anchor="w", pady=5).pack(fill='x', pady=(15, 2))

    def _add_nav_separator(self, parent):
        tk.Frame(parent, height=1, bg="#34495e").pack(fill='x', pady=5)

    def _create_nav_button(self, parent, text, command, state="normal", bg_color=None):
        """Crée un bouton stylisé pour la barre de navigation"""
        if bg_color is None:
            bg_color = self.COLORS["bg_dark"]
            fg_color = self.COLORS["text_white"]
            hover_color = "#34495e"
        else:
            fg_color = "white"
            hover_color = self._adjust_color_lightness(bg_color, 0.9)

        btn = tk.Button(
            parent,
            text=text,
            command=command,
            state=state,
            bg=bg_color,
            fg=fg_color,
            activebackground=hover_color,
            activeforeground=fg_color,
            relief="flat",
            bd=0,
            font=("Segoe UI", 10),
            anchor="w",
            padx=10,
            pady=6,
            cursor="hand2"
        )
        btn.pack(fill='x', pady=1)

        def on_enter(e):
            if btn['state'] != 'disabled':
                # On ne change la couleur au survol que si ce n'est pas le bouton actif (qui est déjà bleu)
                if btn['bg'] != self.COLORS["accent"]:
                     btn['bg'] = hover_color
        
        def on_leave(e):
            if btn['state'] != 'disabled':
                # On ne remet la couleur sombre que si ce n'est pas le bouton actif
                if btn['bg'] != self.COLORS["accent"]:
                    btn['bg'] = bg_color

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn

    def _adjust_color_lightness(self, color_hex, factor):
        return color_hex

    def highlight_module_button(self, active_key):
        """Met en évidence le module actif et réinitialise les autres"""
        module_keys = ['varexp', 'comm', 'event', 'exprv', 'cyclic', 'vartreat']
        
        for key in module_keys:
            if key in self.buttons:
                if key == active_key:
                    self.buttons[key].config(bg=self.COLORS["accent"], fg="white")
                else:
                    self.buttons[key].config(bg=self.COLORS["bg_dark"], fg="white")

    def highlight_button(self, key):
        pass

    # ================= REFRESH TREE (OPTIMISÉ) =================
    def refresh_tree(self, focus_idx=None):
        """
        Rafraichit l'arbre.
        focus_idx : Index (dans self.filtered_indices) sur lequel on veut centrer la vue.
        """
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.visible_columns
    
        for col in self.visible_columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=150, anchor='center', stretch=True)
    
        # === OPTIMISATION FENÊTRAGE ===
        DISPLAY_LIMIT = 1000
        total_filtered = len(self.filtered_indices)
        
        # Calcul de la plage d'affichage (Start - End)
        start_index = 0
        if focus_idx is not None:
            # On essaie de centrer la vue sur l'élément trouvé
            half_window = DISPLAY_LIMIT // 2
            start_index = max(0, focus_idx - half_window)
            
            # Si on est trop près de la fin, on ajuste
            if start_index + DISPLAY_LIMIT > total_filtered:
                start_index = max(0, total_filtered - DISPLAY_LIMIT)
        
        end_index = min(total_filtered, start_index + DISPLAY_LIMIT)
        
        # Sélection du morceau de données à afficher
        indices_to_display = self.filtered_indices[start_index:end_index]
        
        count = start_index
        for real_index in indices_to_display:
            row = self.data[real_index]
            values = []
            for col in self.visible_columns:
                try:
                    col_idx = self.headers.index(col)
                    val = row[col_idx] if col_idx < len(row) else ""
                    values.append(val)
                except ValueError:
                    values.append("")
            
            tag = 'evenrow' if count % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", iid=str(real_index), values=values, tags=(tag,))
            count += 1
    
        # Mise à jour status bar avec info de pagination
        self.update_status_bar(display_info=(start_index, end_index))

    # ================= LOGIQUE MÉTIER ORIGINALE =================
    
    def open_search_replace(self):
        if not self.data:
            return
    
        win = tk.Toplevel(self.root)
        win.title("Recherche / Remplacement")
        win.geometry("500x280")
        win.configure(bg="white")
    
        tk.Label(win, text="Rechercher :", bg="white").pack(anchor='w', padx=10, pady=5)
        search_entry = ttk.Entry(win)
        search_entry.pack(fill='x', padx=10)
    
        tk.Label(win, text="Remplacer par (optionnel) :", bg="white").pack(anchor='w', padx=10, pady=5)
        replace_entry = ttk.Entry(win)
        replace_entry.pack(fill='x', padx=10)
    
        # État de la recherche
        if not hasattr(self, "search_state"):
            self.search_state = {"search": None, "last_index": -1, "filtered_indices": []}
    
        # ---- Recherche suivante ----
        # ---- Recherche suivante (CORRIGÉE) ----
        def search_next():
            search = search_entry.get().lower() # On met tout en minuscule pour la recherche
            if not search:
                return
    
            # Nouvelle recherche ?
            if self.search_state["search"] != search:
                self.search_state["search"] = search
                self.search_state["last_index"] = -1
                self.search_state["filtered_indices"] = self.filtered_indices.copy()
    
            found = False
            start = self.search_state["last_index"] + 1
            
            # Recherche dans les DONNÉES (pas juste l'affichage)
            for pos in range(start, len(self.search_state["filtered_indices"])):
                real_index = self.search_state["filtered_indices"][pos]
                row = self.data[real_index]
                
                # Vérification si le terme existe dans la ligne
                row_str = " ".join([str(x).lower() for x in row])
                if search in row_str:
                    # TROUVÉ !
                    
                    # 1. On "téléporte" l'affichage à cette position (pos)
                    self.refresh_tree(focus_idx=pos)
                    
                    # 2. On sélectionne la ligne
                    if self.tree.exists(str(real_index)):
                        self.tree.see(str(real_index))
                        self.tree.selection_set(str(real_index))
                    
                    self.search_state["last_index"] = pos
                    found = True
                    break
            
            if not found:
                messagebox.showinfo("Recherche", "Fin des occurrences")
                self.search_state["last_index"] = -1
    
        # ---- Remplacer tout ----
        def replace_all():
            search = search_entry.get()
            replace = replace_entry.get()
            if not search:
                return
    
            undo_batch = []
            for real_index in self.filtered_indices:
                row = self.data[real_index]
                new_row = row.copy()
                changed = False
                for i, val in enumerate(row):
                    if search in str(val):
                        new_row[i] = str(val).replace(search, replace)
                        changed = True
                if changed:
                    undo_batch.append((real_index, row.copy()))
                    self.data[real_index] = new_row
    
            if undo_batch:
                self.undo_stack.append(undo_batch)
                self.modified = True
                self.apply_filter()
                messagebox.showinfo("Remplacement", f"{len(undo_batch)} lignes modifiées")
    
        # ---- Boutons ----
        btn_frame = tk.Frame(win, bg="white")
        btn_frame.pack(pady=15)
    
        tk.Button(btn_frame, text="Exécuter / Suivant", bg=self.COLORS["accent"], fg="white", relief="flat", padx=10, command=search_next).pack(side='left', padx=10)
        tk.Button(btn_frame, text="Remplacer tout", bg=self.COLORS["warning"], fg="white", relief="flat", padx=10, command=replace_all).pack(side='left', padx=10)
        tk.Button(btn_frame, text="Fermer", command=win.destroy).pack(side='right', padx=10)
    
    def check_unsaved_changes(self):
        if self.modified:
            res = messagebox.askyesnocancel(
                "Modifications non enregistrées",
                "Vous avez des modifications non enregistrées.\nVoulez-vous enregistrer avant de continuer ?"
            )
            if res is None:  # Annuler
                return False
            elif res:
                self.save_file()
        return True

    # ================= SCROLL =================
    def scroll_top(self):
        if not self.filtered_indices:
            return
            
        # 1. On recharge l'arbre en ciblant le tout début (index 0)
        self.refresh_tree(focus_idx=0)
        
        # 2. On sélectionne visuellement la première ligne
        children = self.tree.get_children()
        if children:
            first_item = children[0]
            self.tree.see(first_item)
            self.tree.selection_set(first_item)

    def scroll_bottom(self):
        if not self.filtered_indices:
            return
            
        # 1. On calcule l'index du dernier élément filtré
        last_idx_in_filter = len(self.filtered_indices) - 1
        
        # 2. On recharge l'arbre en centrant sur la fin
        self.refresh_tree(focus_idx=last_idx_in_filter)
        
        # 3. On sélectionne visuellement la dernière ligne chargée
        children = self.tree.get_children()
        if children:
            last_item = children[-1]
            self.tree.see(last_item)
            self.tree.selection_set(last_item)

    # ================= COPIER/PASTER (BUG FIXED) =================
    def copy_rows(self, event=None):
        widget = self.root.focus_get()
        if widget and widget.winfo_class() in ("Entry", "TEntry"):
            return
    
        self.clipboard_rows = []
        for iid in self.tree.selection():
            real_index = int(iid)
            self.clipboard_rows.append(self.data[real_index].copy())
    
        # [FIX] Suppression de la ligne qui bloquait l'UNDO
        # self.copy_undo_stack_size = len(self.undo_stack) 

    def paste_rows(self, event=None):
        widget = self.root.focus_get()
        if widget and widget.winfo_class() in ("Entry", "TEntry"):
            return
    
        if not self.clipboard_rows:
            return
    
        start_index = len(self.data)
        for row in self.clipboard_rows:
            self.data.append(row.copy())
    
        self.undo_stack.append([
            (start_index + i, row.copy())
            for i, row in enumerate(self.clipboard_rows)
        ])
    
        # [FIX] Plus besoin de gérer copy_undo_stack_size
        # self.copy_undo_stack_size = None
    
        self.modified = True
        self.apply_filter()

    # ================= DELETE (FIXED) =================
    def delete_selected_rows(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
    
        real_indices = sorted([int(iid) for iid in selected])
    
        # Undo
        self.undo_stack.append([(idx, self.data[idx].copy()) for idx in real_indices])
    
        for idx in reversed(real_indices):
            del self.data[idx]
    
        self.apply_filter()
        self.modified = True
        # [FIX] Plus besoin de gérer copy_undo_stack_size
        # self.copy_undo_stack_size = None

    # ================= UNDO (BUG FIXED) =================
    def undo(self, event=None):
        """Annule la dernière action (Supporte modifs unitaires ET snapshots)."""
        if not self.undo_stack:
            if hasattr(self, 'status_var'): self.status_var.set("Rien à annuler.")
            return

        last_action = self.undo_stack.pop()

        # --- CAS 1 : C'EST UN SNAPSHOT (Sauvegarde complète) ---
        # On détecte si c'est une liste de listes (données brutes)
        is_snapshot = False
        if isinstance(last_action, list) and len(last_action) > 0:
            # Si le premier élément est une liste, c'est un tableau de données entier
            if isinstance(last_action[0], list):
                is_snapshot = True
        elif isinstance(last_action, list) and len(last_action) == 0:
            # Cas rare : retour à un tableau vide
            # Pour simplifier, on considère que si ce n'est pas des tuples, c'est un snapshot
            is_snapshot = True

        if is_snapshot:
            # Restauration complète brutale (Rapide pour les gros blocs)
            self.data = last_action
            # On réinitialise les filtres pour éviter des index hors limites
            self.filtered_indices = list(range(len(self.data)))
            self.refresh_tree()
            if hasattr(self, 'status_var'): self.status_var.set("Restauration complète effectuée.")
            return

        # --- CAS 2 : C'EST UNE LISTE D'ACTIONS (Votre logique existante) ---
        # Si ce sont des tuples (row, col, val), on applique votre boucle
        for action in reversed(last_action):
            # Si action est une modification de cellule : 3 valeurs
            if isinstance(action, tuple) and len(action) == 3:
                row_idx, col_idx, old_value = action
                # Sécurité dimensions
                if row_idx < len(self.data):
                    while len(self.data[row_idx]) <= col_idx:
                        self.data[row_idx].append("")
                    self.data[row_idx][col_idx] = old_value

                    # Mise à jour Visuelle unitaire (Optimisation)
                    if str(row_idx) in self.tree.get_children():
                        values = []
                        display_cols = [c for c in self.headers if c in self.visible_columns]
                        for col in display_cols:
                            try:
                                idx_h = self.headers.index(col)
                                val = self.data[row_idx][idx_h] if idx_h < len(self.data[row_idx]) else ""
                                values.append(val)
                            except ValueError:
                                values.append("")
                        self.tree.item(str(row_idx), values=values)

                        # Restauration surlignage
                        col_name = self.headers[col_idx] if col_idx < len(self.headers) else ""
                        if col_name:
                             self.cell_templates.setdefault(str(row_idx), {})[col_name] = old_value

            # Sinon, action classique sur ligne entière : 2 valeurs
            elif isinstance(action, tuple) and len(action) == 2:
                idx, row_data = action
                if row_data is None:
                    # C'était un ajout -> on supprime
                    if idx < len(self.data):
                        del self.data[idx]
                else:
                    # C'était une suppression -> on remet
                    if idx < len(self.data):
                        self.data[idx] = row_data
                    else:
                        self.data.append(row_data)

        # Finitions communes
        self.apply_filter()
        self.modified = True
        if hasattr(self, 'status_var'): self.status_var.set("Modification annulée.")

    # ================= EDIT CELL =================
    def edit_cell(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column:
            return
    
        real_index = int(item)
        col_index = int(column.replace('#', '')) - 1
        
        if col_index < 0 or col_index >= len(self.visible_columns):
            return

        col_name = self.visible_columns[col_index]
        
        try:
            header_idx = self.headers.index(col_name)
        except ValueError:
            return # Colonne non trouvée dans les headers
    
        # Valeur actuelle dans le Treeview
        values = self.tree.item(item, "values")
        if col_index < len(values):
            old_value = values[col_index]
        else:
            old_value = ""
    
        # Fermer l'Entry précédent si existe
        if hasattr(self, "editing_entry") and self.editing_entry:
            self.editing_entry.destroy()
            self.editing_entry = None
    
        # Créer Entry pour édition
        entry = tk.Entry(self.root, font=("Segoe UI", 10), bd=1, relief="solid")
        entry.insert(0, old_value)
        entry.focus()
    
        self.editing_entry = entry
        self.editing_item = (real_index, col_index)
        self.editing_header_idx = header_idx
    
        # Positionner Entry sur la cellule
        bbox = self.tree.bbox(item, f"#{col_index+1}")
        if bbox:
            x, y, w, h = bbox
            entry.place(
                x=self.tree.winfo_rootx() - self.root.winfo_rootx() + x,
                y=self.tree.winfo_rooty() - self.root.winfo_rooty() + y,
                width=w, height=h
            )
    
        # Variable pour suivi dynamique
        var = tk.StringVar(value=old_value)
        entry.config(textvariable=var)
    
        # Valeur originale pour comparaison
        template_val = self.cell_templates.get(item, {}).get(col_name, old_value)
    
        def on_entry_change(*args):
            val = var.get()
            entry.config(bg="#fff6d5" if val != template_val else "white")
    
        var.trace_add("write", on_entry_change)
    
        # Sauvegarde et fermeture Entry
        def save_edit(event=None):
            if not self.editing_entry: return
            
            new_val = var.get()
            entry.destroy()
            self.editing_entry = None
        
            # S'assurer que la ligne a assez de colonnes
            while len(self.data[real_index]) <= header_idx:
                self.data[real_index].append("")
        
            # ====== UNDO STACK ======
            self.undo_stack.append([(real_index, header_idx, self.data[real_index][header_idx])])
        
            # ====== Mise à jour self.data ======
            self.data[real_index][header_idx] = new_val
            
            # ====== Mise à jour interface ======
            # On met à jour directement l'item treeview sans tout recharger
            current_values = list(self.tree.item(item, "values"))
            if col_index < len(current_values):
                current_values[col_index] = new_val
                self.tree.item(item, values=current_values)
        
            # ====== Mise à jour valeur originale pour surlignage futur =====
            self.cell_templates.setdefault(item, {})[col_name] = new_val

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

        
    def sort_by_column(self, col_name):
        idx = self.headers.index(col_name)
    
        # Toggle asc / desc
        asc = self.sort_state.get(col_name, True)
        self.sort_state = {col_name: not asc}
    
        def sort_key(real_index):
            row = self.data[real_index]
            if idx >= len(row):
                return (1, "")
        
            val = row[idx]
        
            if val is None:
                return (1, "")
        
            val = str(val).strip()
            if val == "":
                return (1, "")
        
            try:
                return (0, float(val))
            except ValueError:
                return (1, val.lower())
    
        self.filtered_indices.sort(key=sort_key, reverse=asc)
        self.refresh_tree()

    def update_status_bar(self, display_info=None):
        total = len(self.data)
        visible_total = len(self.filtered_indices)
        
        # Gestion du message de pagination
        range_msg = ""
        if display_info:
            start, end = display_info
            range_msg = f"[Vue: {start+1}-{end}]"
        elif visible_total > 1000:
            range_msg = f"[Vue: 1-1000]"
            
        selection = self.tree.selection()
        line_text = "-"
        if selection:
            try:
                # On essaie de récupérer l'index réel
                line_text = f"ID: {selection[0]}"
            except:
                pass
    
        filter_active = any([self.filter_entry1.get().strip(), self.filter_entry2.get().strip(), self.filter_entry3.get().strip()])
        filter_text = "ACTIF" if filter_active else "Aucun"
        modified_text = "⚠️ Modifié" if self.modified else "Sync"
    
        text = (
            f"Données : {visible_total} / {total} {range_msg} | "
            f"{line_text} | "
            f"Filtre : {filter_text} | "
            f"{modified_text}"
        )
        self.status_var.set(text)

    # ================= FICHIERS (LOAD/SAVE) =================
    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder = folder
            messagebox.showinfo("Dossier sélectionné", f"Dossier choisi : {folder}")
            # Folder n'est pas un module, donc on ne highlight pas de module spécifique
            # mais on peut highlighter le bouton dossier si on veut
            self.buttons['folder'].config(bg=self.COLORS["accent"])

    def _load_file_generic(self, path=None, force_headers=None, skip_first_line=False, button_key=None):
        if not path:
            path = filedialog.askopenfilename(filetypes=[("DAT files", "*.dat")])
        if not path:
            return
        self.data = []
        self.headers = []
        self.first_line = None
        try:
            with open(path, 'r', encoding='latin-1') as f:
                reader = csv.reader(f, delimiter=',', quotechar='"')
                if skip_first_line:
                    try:
                        self.first_line = next(reader)
                    except StopIteration:
                        self.first_line = None
                for row in reader:
                    self.data.append(row)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier : {e}")
            return
        if force_headers:
            self.headers = force_headers
        else:
            for row in self.data:
                if any(any(c.isalpha() for c in cell) for cell in row if isinstance(cell, str)):
                    self.headers = row
                    break
        self.visible_columns = [h for h in self.headers if h and "unnamed" not in h.lower()]
        self.filtered_indices = list(range(len(self.data)))
        
        # Initialisation du last_tagname à partir du fichier chargé
        col_tag = self.find_header("Tagname")
        if col_tag:
            idx = self.headers.index(col_tag)
            self.last_tagname = 0
            for row in reversed(self.data):
                try:
                    val = int(row[idx])
                    self.last_tagname = val
                    break
                except (ValueError, IndexError):
                    continue
        for combobox in [self.column_filter1, self.column_filter2, self.column_filter3]:
            combobox['values'] = self.visible_columns
            if self.visible_columns:
                combobox.current(0)
        
        # Refresh UI
        self.refresh_tree()

        # Highlight logic for Modules
        if button_key:
            self.highlight_module_button(button_key)
            
        self.modified = False

    def save_file(self):
        if not self.data:
            return
        path = filedialog.asksaveasfilename(defaultextension=".dat",
                                            filetypes=[("DAT files", "*.dat")])
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='latin-1') as f:
                writer = csv.writer(f, delimiter=',', quotechar='"')
                if self.first_line:
                    writer.writerow(self.first_line)
                writer.writerows(self.data)
            messagebox.showinfo("Succès", "Fichier enregistré")
            self.modified = False
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def close_all_popups(self):
        """
        Ferme les fenêtres contextuelles (Recherche, Colonnes) 
        MAIS laisse la fenêtre de Comparaison ouverte.
        """
        # 1. Fermer la fenêtre de Recherche
        if hasattr(self, 'search_window') and self.search_window is not None:
            try:
                self.search_window.destroy()
            except:
                pass
            self.search_window = None

        # 2. Fermer la fenêtre de Colonnes (si vous en avez une gérée ainsi)
        if hasattr(self, 'column_window') and self.column_window is not None:
            try:
                self.column_window.destroy()
            except:
                pass
            self.column_window = None
            
    def open_any_dat_file(self):
        path = filedialog.askopenfilename(
            title="Ouvrir un fichier .DAT",
            filetypes=[("Fichiers DAT", "*.dat"), ("Tous les fichiers", "*.*")]
        )
        if not path:
            return
    
        self.selected_folder = os.path.dirname(path)
    
        import csv
    
        data = []
        try:
            with open(path, 'r', encoding='latin-1') as f:
                reader = csv.reader(f, delimiter=',', quotechar='"')
                for row in reader:
                    data.append(row)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier : {e}")
            return
    
        if not data:
            messagebox.showwarning("Attention", "Le fichier est vide.")
            return
    
        # 🔥 Longueur max
        max_cols = max(len(r) for r in data)
    
        # 🔥 Normalisation de toutes les lignes
        for row in data:
            if len(row) < max_cols:
                row.extend([""] * (max_cols - len(row)))
    
        # 🔥 Génération automatique des headers
        self.headers = [f"Col_{i+1}" for i in range(max_cols)]
    
        self.data = data
        self.visible_columns = self.headers.copy()
        self.filtered_indices = list(range(len(self.data)))
    
        # Highlight special
        self.highlight_module_button('open_any')
        self.buttons['open_any'].config(bg=self.COLORS["accent"])

        self.buttons['create_var'].config(state="disabled")
        self.buttons['create_event'].config(state="disabled")       # active création EVENT
        self.buttons['create_exprv'].config(state="disabled")     # désactive création EXPRV
        self.buttons['create_cyclic'].config(state="disabled")    # désactive création CYCLIC
        self.buttons['create_vartreat'].config(state="disabled")
    
        self.modified = False
        
        # Mettre à jour les colonnes visibles et filtrables
        self.visible_columns = [h for h in self.headers if h and "unnamed" not in h.lower()]
        self.filtered_indices = list(range(len(self.data)))
        
        # 🔹 Mettre à jour les combobox de filtrage pour le nouveau fichier
        for combobox in [self.column_filter1, self.column_filter2, self.column_filter3]:
            combobox['values'] = self.visible_columns
            if self.visible_columns:
                combobox.current(0)
        
        self.refresh_tree()

    # ================= SCROLLING FILES =================
    def load_varexp(self):
        self.close_all_popups()
        if not self.selected_folder:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'abord")
            return
        if not self.check_unsaved_changes():
            return
        path = os.path.join(self.selected_folder, "varexp.dat")
        if not os.path.exists(path):
            messagebox.showerror("Erreur", f"Fichier varexp.dat non trouvé dans {self.selected_folder}")
            return
        self._load_file_generic(path=path, skip_first_line=True, button_key='varexp')
        self.buttons['create_var'].config(state="normal")
        self.buttons['create_event'].config(state="disabled")       # active création EVENT
        self.buttons['create_exprv'].config(state="disabled")     # désactive création EXPRV
        self.buttons['create_cyclic'].config(state="disabled")    # désactive création CYCLIC
        self.buttons['create_vartreat'].config(state="disabled")

    def load_comm(self):
        self.close_all_popups()
        if not self.selected_folder:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'abord")
            return
        if not self.check_unsaved_changes():
            return
        path = os.path.join(self.selected_folder, "COMM.DAT")
        if not os.path.exists(path):
            messagebox.showerror("Erreur", f"Fichier COMM.DAT non trouvé dans {self.selected_folder}")
            return
        self._load_file_generic(path=path, force_headers=self.COMM_DEFAULT_HEADERS, skip_first_line=True, button_key='comm')
        self.buttons['create_var'].config(state="disabled")       # désactive création VAREXP
        self.buttons['create_event'].config(state="disabled")       # active création EVENT
        self.buttons['create_exprv'].config(state="disabled")     # désactive création EXPRV
        self.buttons['create_cyclic'].config(state="disabled")    # désactive création CYCLIC
        self.buttons['create_vartreat'].config(state="disabled")

    def load_event(self):
        self.close_all_popups()
        if not self.selected_folder:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'abord")
            return
        if not self.check_unsaved_changes():
            return
        path = os.path.join(self.selected_folder, "EVENT.DAT")
        if not os.path.exists(path):
            messagebox.showerror("Erreur", f"Fichier EVENT.DAT non trouvé dans {self.selected_folder}")
            return
        self._load_file_generic(path=path, force_headers=self.EVENT_DEFAULT_HEADERS, button_key='event')
        self.buttons['create_var'].config(state="disabled")       # désactive création VAREXP
        self.buttons['create_event'].config(state="normal")       # active création EVENT
        self.buttons['create_exprv'].config(state="disabled")     # désactive création EXPRV
        self.buttons['create_cyclic'].config(state="disabled")    # désactive création CYCLIC
        self.buttons['create_vartreat'].config(state="disabled")

    def load_exprv(self):
        self.close_all_popups()
        if not self.selected_folder:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'abord")
            return
        if not self.check_unsaved_changes():
            return
        path = os.path.join(self.selected_folder, "Exprv.DAT")
        if not os.path.exists(path):
            messagebox.showerror("Erreur", f"Fichier Exprv.DAT non trouvé dans {self.selected_folder}")
            return
        self._load_file_generic(path=path, force_headers=self.EXPRV_DEFAULT_HEADERS, button_key='exprv')
        self.buttons['create_var'].config(state="disabled")
        self.buttons['create_event'].config(state="disabled")
        self.buttons['create_exprv'].config(state="normal")
        self.buttons['create_cyclic'].config(state="disabled")
        self.buttons['create_vartreat'].config(state="disabled")

    def load_cyclic(self):
        self.close_all_popups()
        if not self.selected_folder:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'abord")
            return
        if not self.check_unsaved_changes():
            return
        path = os.path.join(self.selected_folder, "CYCLIC.DAT")
        if not os.path.exists(path):
            messagebox.showerror("Erreur", f"Fichier CYCLIC.DAT non trouvé dans {self.selected_folder}")
            return
        self._load_file_generic(path=path, force_headers=self.CYCLIC_DEFAULT_HEADERS, button_key='cyclic')
        self.buttons['create_var'].config(state="disabled")
        self.buttons['create_event'].config(state="disabled")
        self.buttons['create_exprv'].config(state="disabled")
        self.buttons['create_cyclic'].config(state="normal")
        self.buttons['create_vartreat'].config(state="disabled")

    def load_vartreat(self):
        self.close_all_popups()
        if not self.selected_folder:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un dossier d'abord")
            return
        if not self.check_unsaved_changes():
            return
        path = os.path.join(self.selected_folder, "VARTREAT.DAT")
        if not os.path.exists(path):
            messagebox.showerror("Erreur", f"Fichier VARTREAT.DAT non trouvé dans {self.selected_folder}")
            return
        self._load_file_generic(path=path, force_headers=self.VARTREAT_DEFAULT_HEADERS, button_key='vartreat')
        self.buttons['create_var'].config(state="disabled")
        self.buttons['create_event'].config(state="disabled")
        self.buttons['create_exprv'].config(state="disabled")
        self.buttons['create_cyclic'].config(state="disabled")
        self.buttons['create_vartreat'].config(state="normal")
        
    def find_header(self, name):
        for h in self.headers:
            if h.lower() == name.lower():
                return h
        return None
    
    # ================= CREATION EVENT / EXPRV / CYCLIC =================
    def open_create_generic(self, filetype):
        if filetype not in ["event", "exprv", "cyclic", "vartreat"]:
            return
    
        win = tk.Toplevel(self.root)
        win.title(f"Création {filetype.upper()}")
        win.configure(bg="white")
        if filetype == 'vartreat' :
            win.geometry("750x650")
        else :
            win.geometry("550x550")
    
        # Désactiver les autres boutons create_
        for key, btn in self.buttons.items():
            if key.startswith("create_") and key != f"create_{filetype}":
                btn.config(state="disabled")
    
        default_headers = {
            "event": self.EVENT_DEFAULT_HEADERS,
            "exprv": self.EXPRV_DEFAULT_HEADERS,
            "cyclic": self.CYCLIC_DEFAULT_HEADERS,
            "vartreat": self.VARTREAT_DEFAULT_HEADERS
        }
        headers = default_headers[filetype]
    
        def default_value_for_title(title):
            t = str(title).strip().lower()
            if t in {"=0", "0", "00"}:
                return "0"
            if t in {"=1", "1"}:
                return "1"
            if t == "vide":
                return ""
            return ""
    
        # Récupération des Noms existants
        name_col = self.find_header("Nom")
        existing_names = []
        if name_col:
            name_idx = self.headers.index(name_col)
            existing_names = sorted({
                row[name_idx]
                for row in self.data
                if len(row) > name_idx and row[name_idx].strip()
            })
    
        # Frame scrollable
        container = tk.Frame(win, bg="white")
        container.pack(fill="both", expand=True, padx=5, pady=5)
    
        canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="white")
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
        # 🔽 FILTRAGE + COMBOBOX NOM EXISTANT
        tk.Label(scroll_frame, text="Filtrer les configurations existantes", bg="white").grid(
            row=0, column=0, sticky="w", padx=5, pady=(5, 2)
        )
        filter_var = tk.StringVar()
        filter_entry = ttk.Entry(scroll_frame, textvariable=filter_var)
        filter_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=(5, 2))
        tk.Label(scroll_frame, text="Copie d'une configuration existante (Optionnelle)",
                 font=("Segoe UI", 10, "bold"), bg="white").grid(row=1, column=0, sticky="w", padx=5, pady=(2, 5))
        existing_name_cb = ttk.Combobox(scroll_frame, values=existing_names, state="readonly")
        existing_name_cb.grid(row=1, column=1, sticky="ew", padx=5, pady=(2, 5))
        scroll_frame.grid_columnconfigure(1, weight=1)
    
        def update_combobox(*args):
            text = filter_var.get().lower().strip()
            if not text:
                existing_name_cb["values"] = existing_names
            else:
                existing_name_cb["values"] = [n for n in existing_names if text in n.lower()]
    
        filter_var.trace_add("write", update_combobox)
    
        # 🧾 CHAMPS DE SAISIE
        entries = {}
        for i, col in enumerate(headers, start=2):
            if filetype == 'vartreat' :
                tk.Label(scroll_frame, text=col, width=65, anchor="w", bg="white").grid(
                row=i, column=0, sticky="w", padx=5, pady=2
            )
            else :
                tk.Label(scroll_frame, text=col, width=28, anchor="w", bg="white").grid(
                row=i, column=0, sticky="w", padx=5, pady=2
            )
            if filetype == 'vartreat':
                e = ttk.Entry(scroll_frame, width=30)
            else : 
                e = ttk.Entry(scroll_frame)
            e.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            e.insert(0, default_value_for_title(col))
            entries[col] = e
    
        # 🔁 CHARGEMENT DEPUIS UN NOM EXISTANT
        def load_from_existing(event=None):
            selected_name = existing_name_cb.get()
            if not selected_name or not name_col:
                return
            name_idx = self.headers.index(name_col)
            model_row = None
            for row in self.data:
                if len(row) > name_idx and row[name_idx] == selected_name:
                    model_row = row
                    break
            if not model_row:
                return
            for col, entry in entries.items():
                if col in self.headers:
                    idx = self.headers.index(col)
                    val = model_row[idx] if idx < len(model_row) else ""
                    entry.delete(0, tk.END)
                    entry.insert(0, val)
    
        existing_name_cb.bind("<<ComboboxSelected>>", load_from_existing)
    
        # ♻️ RÉINITIALISATION DES CHAMPS
        def reset_fields():
            existing_name_cb.set("")
            filter_var.set("")
            for col, entry in entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, default_value_for_title(col))
    
        # ➕ CRÉATION DE LA LIGNE
        def create_row():
            new_row = [""] * len(self.headers)
            for col, entry in entries.items():
                if col in self.headers:
                    idx = self.headers.index(col)
                    new_row[idx] = entry.get().strip()
    
            # Undo
            self.undo_stack.append([(len(self.data), None)])
    
            # Ajout data
            self.data.append(new_row)
            self.modified = True
    
            # Update View
            self.apply_filter() # Ensure filters are respected
            self.scroll_bottom()
    
        # 🎛️ BOUTONS
        btn_frame = tk.Frame(win, bg="white")
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Réinitialiser", command=reset_fields, bg="#ecf0f1", relief="flat", width=15).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Créer", command=create_row, bg=self.COLORS["success"], fg="white", relief="flat", width=15,
                  font=("Segoe UI", 11, "bold")).pack(side="left", padx=5)
    
    def open_duplicate_branch_window(self):
        """Ouvre une fenêtre pour définir la source et la destination de la duplication."""
        win = tk.Toplevel(self.root)
        win.title("Duplication de Branche")
        win.geometry("600x750")
        win.config(bg="#ecf0f1")
        
        # --- Zone Source ---
        lbl_src = tk.LabelFrame(win, text="1. Branche SOURCE (Ce qu'il faut copier)", bg="#ecf0f1", font=("Segoe UI", 10, "bold"), fg="#c0392b")
        lbl_src.pack(fill="x", padx=10, pady=10)
        
        src_entries = []
        frame_src_grid = tk.Frame(lbl_src, bg="#ecf0f1")
        frame_src_grid.pack(fill="x", padx=5, pady=5)
        
        for i in range(11): # n1 à n11
            tk.Label(frame_src_grid, text=f"n{i+1}", bg="#ecf0f1", width=4).grid(row=i, column=0, padx=2, pady=2)
            e = tk.Entry(frame_src_grid, width=40)
            e.grid(row=i, column=1, padx=2, pady=2)
            src_entries.append(e)

        # --- Zone Destination ---
        lbl_dst = tk.LabelFrame(win, text="2. Branche DESTINATION (Où coller)", bg="#ecf0f1", font=("Segoe UI", 10, "bold"), fg="#27ae60")
        lbl_dst.pack(fill="x", padx=10, pady=10)
        
        dst_entries = []
        frame_dst_grid = tk.Frame(lbl_dst, bg="#ecf0f1")
        frame_dst_grid.pack(fill="x", padx=5, pady=5)
        
        for i in range(11): # n1 à n11
            tk.Label(frame_dst_grid, text=f"n{i+1}", bg="#ecf0f1", width=4).grid(row=i, column=0, padx=2, pady=2)
            e = tk.Entry(frame_dst_grid, width=40)
            e.grid(row=i, column=1, padx=2, pady=2)
            dst_entries.append(e)
            
        # --- Bouton Action ---
        btn_action = tk.Button(win, text="EXÉCUTER LA DUPLICATION", 
                               bg="#3498db", fg="white", font=("Segoe UI", 12, "bold"),
                               command=lambda: self.perform_branch_duplication(src_entries, dst_entries, win))
        btn_action.pack(pady=20, ipadx=10, ipady=5)
        
        tk.Label(win, text="Note: Les variables copiées seront ajoutées à la fin du fichier.", bg="#ecf0f1", fg="#7f8c8d").pack()

    def get_next_tag_id(self):
        """Calcule le prochain TagName disponible (Max + 1) sur tout le fichier."""
        # 1. Trouver la colonne TagName
        tag_col_index = -1
        for idx, h in enumerate(self.headers):
            if h.lower() == "tagname":
                tag_col_index = idx
                break
        
        if tag_col_index == -1:
            return 1 # Pas de colonne TagName, on commence à 1
            
        # 2. Trouver le Max absolu
        current_max = 0
        for row in self.data:
            if tag_col_index < len(row):
                val = str(row[tag_col_index]).strip()
                if val.isdigit():
                    val_int = int(val)
                    # MODIFICATION : On a retiré la limite < 65535.
                    # On prend le max absolu trouvé.
                    if val_int > current_max:
                        current_max = val_int
                        
        return current_max + 1
    
    def perform_branch_duplication(self, src_widgets, dst_widgets, window):
        """
        Logique de copie intelligente (Flexible sur les colonnes + TagName illimité).
        """
        # 1. Récupération des saisies utilisateur
        src_path = [e.get().strip() for e in src_widgets]
        dst_path = [e.get().strip() for e in dst_widgets]
        
        while src_path and src_path[-1] == "": src_path.pop()
        while dst_path and dst_path[-1] == "": dst_path.pop()
        
        if not src_path:
            messagebox.showwarning("Attention", "La branche source est vide. Veuillez spécifier au moins n1.")
            return

        # 2. Identification des colonnes n1..n11 (FLEXIBLE)
        n_indices = []
        columns_found_count = 0
        
        for i in range(1, 12): 
            found_idx = -1
            possible_names = [f"n{i}", f"Chemin n{i}", f"N{i}", f"Level {i}"]
            for name in possible_names:
                if name in self.headers:
                    found_idx = self.headers.index(name)
                    columns_found_count += 1
                    break
            n_indices.append(found_idx)

        if columns_found_count == 0:
            messagebox.showerror("Erreur", "Aucune colonne de niveau (n1, n2...) trouvée.")
            return

        # --- GESTION DU TAGNAME VIA LA NOUVELLE FONCTION ---
        # On calcule le prochain ID une seule fois au début
        next_tag_id = self.get_next_tag_id()
        
        # On retrouve l'index de la colonne pour l'écriture
        tag_col_index = -1
        for idx, h in enumerate(self.headers):
            if h.lower() == "tagname":
                tag_col_index = idx
                break

        # 3. Parcours et Copie
        new_rows = []
        count_copied = 0
        src_len = len(src_path)
        
        for row in self.data:
            # A. Reconstruction chemin
            row_path = []
            for idx in n_indices:
                if idx != -1 and idx < len(row):
                    row_path.append(str(row[idx]).strip())
                else:
                    row_path.append("")
            
            # B. Match Source
            is_match = True
            for k in range(src_len):
                if k < len(row_path):
                    if row_path[k] != src_path[k]:
                        is_match = False
                        break
                else:
                    is_match = False
                    break
            
            if is_match:
                new_row = list(row) 
                
                # C. Nouveau chemin
                suffix = row_path[src_len:]
                final_path = dst_path + suffix
                
                for k, col_idx in enumerate(n_indices):
                    if col_idx != -1:
                        if k < len(final_path):
                            while len(new_row) <= col_idx: new_row.append("")
                            new_row[col_idx] = final_path[k]
                        else:
                            if col_idx < len(new_row): new_row[col_idx] = ""

                # D. Nouveau TagName
                if tag_col_index != -1:
                    while len(new_row) <= tag_col_index: new_row.append("")
                    new_row[tag_col_index] = str(next_tag_id)
                    next_tag_id += 1 # On incrémente pour la prochaine ligne copiée
                
                new_rows.append(new_row)
                count_copied += 1

        # 4. Sauvegarde et Scroll
        if count_copied > 0:
            start_index = len(self.data)
            self.data.extend(new_rows)
            end_index = len(self.data)
            
            new_indices = list(range(start_index, end_index))
            self.filtered_indices.extend(new_indices)
            
            self.refresh_tree()
            
            last_item_id = str(end_index - 1)
            try:
                self.tree.see(last_item_id)
                self.tree.selection_set(last_item_id)
            except: pass

            msg = f"{count_copied} variables dupliquées (IDs démarrant à {next_tag_id - count_copied})."
            messagebox.showinfo("Succès", msg)
            window.destroy()
        else:
            messagebox.showwarning("Résultat", "Aucune variable trouvée pour cette branche source.")
            
    def open_create_variable(self):
        win = tk.Toplevel(self.root)
        win.title("Création de variable")
        win.geometry("600x550")
        win.configure(bg="white")
    
        # Case à cocher pour sauvegarde des paramètres avancés
        save_adv_var = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(
            win, 
            text="Sauvegarde des paramètres avancés pour la prochaine création de variable", 
            variable=save_adv_var,
            wraplength=680,
            justify="left",
            bg="white"
        )
        cb.pack(pady=5, anchor="w")
    
        # ---------- Classe ----------
        tk.Label(win, text="Classe", bg="white").pack(anchor="w", padx=10)
        class_cb = ttk.Combobox(win, values=list(self.VAREXP_TEMPLATES.keys()), state="readonly")
        class_cb.pack(fill="x", padx=10)
        class_cb.current(0)
        
        # Label pour afficher la description
        desc_label = tk.Label(win, text=DatEditor.VAREXP_DESCRIPTIONS.get(class_cb.get(), ""), wraplength=480, fg=self.COLORS["accent"], bg="white")
        desc_label.pack(anchor="w", padx=10, pady=2)
        
        # Mettre à jour la description quand l'utilisateur change de classe
        def update_class_desc(event=None):
            desc_label.config(text=DatEditor.VAREXP_DESCRIPTIONS.get(class_cb.get(), ""))
        
        class_cb.bind("<<ComboboxSelected>>", update_class_desc)
        
        # ---------- Nom et chemin ----------
        tk.Label(win, text="Nom de la variable", bg="white").pack(anchor="w", padx=10, pady=(10, 0))
        name_entry = ttk.Entry(win)
        name_entry.pack(fill="x", padx=10)
    
        tk.Label(win, text="Chemin (1 élément par ligne – max 11 éléments)", bg="white").pack(anchor="w", padx=10, pady=(10,0))
        path_entries = []
        for i in range(11):
            f = tk.Frame(win, bg="white")
            f.pack(fill="x", padx=10)
            tk.Label(f, text=f"n{i+1}", width=3, anchor="w", bg="white", fg="#95a5a6").pack(side='left')
            e = ttk.Entry(f)
            e.pack(side='left', fill="x", expand=True)
            path_entries.append(e)
    
        # ---------- Paramètres avancés ----------
        advanced_params = {}  # { "NomColonne": Entry widget }
    
        categories = {
            "Informations générales variables": ["Description", "DescriptionAlt", "Domain", "Nature", 
                                                 "AssociatedLabels", "Inhibited", "Simulated", "Saved", "Broadcast", "PermanentScan",
                                                 "Recorder", "Log0_1", "Log1_0", "WithInitialValue", "InitialValue",
                                                 "ServerListName", "ClientListName", "Source", "BrowsingLevel",
                                                 "AlarmAcknowledgmentLevel", "AlarmMaskLevel", 
                                                 "AlarmMaintenanceLevel", "DeadbandType", 
                                                 "MaskExpressionBranch", "MaskExpressionTemplate",
                                                 "MessageAlarm"],
            "Equipement": ["Eqt_NetworkName", "Eqt_EqtName", "Eqt_FrameName", "Eqt_Type", "Eqt_Index", "Eqt_IndexComp", "Eqt_SizeBit", "Eqt_TextEncoding"],
            "Attributs Etendus": ["ExtBinary", "ExtText3", "ExtText4", "ExtText5", "ExtText6", "ExtText7", "ExtText8",
                                  "ExtText9", "ExtText10", "ExtText11", "ExtText12", "ExtText13", "ExtText14", "ExtText15", "ExtText16"],
            "Commandes et alarmes": ["BitCommandLevel", "AlarmLevel", "AlarmActiveAt1", 
                                     "MaskDependency", "AlarmTemporization"],
            "Mesures": ["Unit", "DeadbandValue", "MinimumValue", "MaximumValue", "ScaledValue",
                        "DeviceMinimumValue", "DeviceMaximumValue", "Format", "VariableMinimumValue",
                        "VariableMaximumValue", "ControlMinimumValue", "ControlMaximumValue",
                        "RegisterCommandLevel", "VariableControlMinimumValue", "VariableControlMaximumValue"],
            "Compteurs": ["Counter_StepSize", "Counter_Type", "Counter_CountBitName", "Counter_CountBitTransition", "Counter_ResetBitName", "Counter_ResetBitTransition"],
            "Chrono": ["Chrono_Period", "Chrono_Type", "Chrono_EnableBitName", "Chrono_EnableBitTransition", "Chrono_ResetBitName", "Chrono_ResetBitTransition"],
            "Valeurs de seuils sur mesure": ["ThresholdHysterisis", "ThresholdValue", "ThresholdHigh", "ThresholdSource", "ThresholdSystem", "ThresholdTypeInSystem"],
            "Texte": ["TextSize", "TextCommandLevel"],
            "Communication LNS": ["LNS_NetworkAlias", "LNS_NodeAlias", "LNS_NvName", "LNS_PollingPeriod", "LNS_DefaultMonitoring", "LNS_FieldName"],
            "Communication DDE": ["DDE_Server", "DDE_Item", "DDE_Format", "DDE_RangBit", "DDE_AutoItemName", "DDE_Label"],
            "Communication OPC": ["OPC_Server", "OPC_Group", "OPC_ItemId", "OPC_AccessPath", "OPC_IsArray", "OPC_ArrayIndex", "OPC_CustomizationExpression"],
            "Communication OPCUA": ["OPCUA_NetworkName", "OPCUA_ClientName", "OPCUA_MonitoringName", "OPCUA_Identifier", "OPCUA_IdentifierType", "OPCUA_NamespaceIndex"],
            "Communication SNMP": ["SNMP_NetworkName", "SNMP_DeviceName", "SNMP_ScanGroupName", "SNMP_DataType", "SNMP_OID", "SNMP_DisableReading", 
                                   "SNMP_WithInitialValue", "SNMP_InitialValue", "SNMP_Offset", "SNMP_ExtractionField", "SNMP_RemoveNoPrintableCharacters"],
            "Communication BACnet": ["BACnet_NetworkName", "BACnet_DeviceName", "BACnet_ObjectType", "BACnet_ObjectInstance", "BACnet_Property", "BACnet_Fields", 
                                     "BACnet_PollingPeriod", "BACnet_MonitoringType", "BACnet_Priority", "BACnet_Type", "BACnet_VarType", "BACnet_AlarmType"],
            "Communication IEC61850": ["IEC61850_MappingType", "IEC61850_Network", "IEC61850_PhysicalDevice", "IEC61850_DataGroup", "IEC61850_DataGroupMember", 
                                       "IEC61850_DataGroupMemberField", "IEC61850_CommonDataClass", "IEC61850_ControlModel", "IEC61850_NotUseTimeStamp", "IEC61850_NotUseQuality"],
            "Communication 104": ["104_NetworkName", "104_DeviceName", "104_SectorName", "104_IOA", "104_Type", "104_WriteIOA", "104_SBO", "104_QLorQU", 
                                  "104_WriteTimeTag", "104_MappingBit", "104_ReadTimeTag"]
        }
        
        CATEGORY_DESCRIPTIONS = {
                "Description": "Description de la variable",
                "Inhibited": "Variable inhibée (I ou vide)",
                "Simulated": "Variable simulée (S ou vide)",
                "Saved": "Variable sauvegardée (P ou vide)",
                "Broadcast": "Accès distant (0 ou 1)",
                "PermanentScan" : "Scrutation permanente pour synoptique (tous postes = 0, poste serveur = 2, aucun = 1)",
                "Recorder": "Variable magnétorisée (1 ou 0)", 
                "Log0_1": "Variable consignée de 0 vers 1 (1 ou 0)", 
                "Log1_0": "Variable consignée de 1 vers 0 (1 ou 0)", 
                "WithInitialValue": "Variable avec valeur initiale (1 ou 0)",
                "InitialValue": "Valeur initiale (float)",
                "ServerListName": "Nom de la liste serveur", 
                "ClientListName": "Nom de la liste client", 
                "Source": "Source de la variable (Interne = I, OPCUA = U, SNMP = S, Equipement = E, ...)", 
                "BrowsingLevel": "Niveau de recherche (entre 0 et 29)",
                "AlarmAcknowledgmentLevel": "Niveau d'acquittement (entre 0 et 29)", 
                "AlarmMaskLevel": "Niveau de masquage (entre 0 et 29)", 
                "AlarmMaintenanceLevel": "Niveau de prise en compte (entre 0 et 29)", 
                "Eqt_NetworkName": "Nom du réseau", 
                "Eqt_EqtName": "Nom de l'équipement", 
                "Eqt_FrameName": "Nom de la trame", 
                "Eqt_Type": "Type de la trame (bit = B, mot = M, I, U)", 
                "Eqt_Index": "Offset octet", 
                "Eqt_IndexComp": "Offset bit", 
                "Eqt_SizeBit": "Taille de la varible en bit (1 si B, 32 si M, 16 si I ou U)", 
                "ExtBinary": "Enable les attributs étendus (0 ou vide)", 
                "AlarmLevel": "Priorité (entre 0 et 29)", 
                "AlarmActiveAt1": "Déclenchement Positive ou Négative (1 ou 0)", 
                "AlarmTemporization": "Temporisation en secondes des alarmes (int)",
                "Unit": "Unité de la mesure", 
                "MinimumValue": "Valeur min", 
                "MaximumValue": "Valeur max", 
                "Format": "Format du rendu de la mesure", 
                "ControlMinimumValue": "Valeur minimale de commande de registre", 
                "ControlMaximumValue": "Valeur maximale de commande de registre",
                "RegisterCommandLevel": "Niveau de commande de registre", 
                "Counter_StepSize": "Valeur de pas (int)", 
                "Counter_Type": "Type de compteur : Décrémental ou Incrémental (0 ou 1)", 
                "Counter_CountBitName": "Variable état associé (nom)", 
                "Counter_CountBitTransition": "Enclenchement de l'état associé (à 0 ou à 1)", 
                "Counter_ResetBitName": "Variable bit d'initialisation (nom)", 
                "Counter_ResetBitTransition": "Enclenchement du bit d'initialisation (à 0 ou à 1)",
                "Chrono_Period": "Période d'incrémentation du chrono (100 = 1 sec, 6000 = 1 min, ...)", 
                "Chrono_Type": "Type du chrono (1)", 
                "Chrono_EnableBitName": "Variable de déclenchement (nom)", 
                "Chrono_EnableBitTransition": "Enclenchement du chrono sur la variable de déclenchement à 1 ou 0 (1 ou 0)", 
                "Chrono_ResetBitName": "Variable d'initialisation (nom)", 
                "Chrono_ResetBitTransition": "Initialisation du chrono sur la variable d'initialisation à 1 ou 0 (1 ou 0)",
                "ThresholdHysterisis": "Hysteresis (float)", 
                "ThresholdValue": "Valeur de seuil (float)", 
                "ThresholdHigh": "Seuil haut (1 ou 0)", 
                "ThresholdSource": "Variable reliée au seuil", 
                "ThresholdSystem": "Type de seuil (ppphaut|pphaut|phaut|haut = 0, pphaut|phaut|haut|bas = 1, phaut|haut|bas|pbas = 2, haut|bas|pbas|ppbas = 3 sinon 4)", 
                "ThresholdTypeInSystem": "Type de variable seuil (de 0 à 3 du seuil le plus haut à celui le plus bas)",
                "TextSize": "Taille maximum de la chaîne de caractère en octets (int)", 
                "TextCommandLevel": "Niveau de commande (entre 0 et 29)",
                "OPCUA_NetworkName": "Nom du réseau OPCUA", 
                "OPCUA_ClientName": "Nom du client OPCUA", 
                "OPCUA_MonitoringName": "Nom groupe de scrutation", 
                "OPCUA_Identifier": "Identificateur de la variable sur le serveur OPCUA", 
                "SNMP_NetworkName": "Nom du réseau SNMP", 
                "SNMP_DeviceName": "Nom de l'équipement", 
                "SNMP_ScanGroupName": "Nom du groupe de scrutation", 
                "SNMP_DataType": "Type de données SNMP", 
                "SNMP_OID": "OID SNMP", 
                "SNMP_DisableReading": "Désactivation lecture (0 ou 1)", 
                "SNMP_WithInitialValue": "SNMP avec valeur initiale (0 ou 1)", 
                "SNMP_InitialValue": "Valeur Initiale (int ou vide)", 
                "SNMP_RemoveNoPrintableCharacters": "Suppression des caractère spéciaux (0 ou 1)",
                }

        def open_advanced():
            EXCLUDED_COLUMNS = {"Class", "Tagname", "Nom"} | {f"n{i}" for i in range(1, 13)}
        
            adv_win = tk.Toplevel(win)
            adv_win.title("Paramètres avancés")
            adv_win.geometry("800x800")
            adv_win.configure(bg="white")
        
            # ===================== TAGNAME MODELE =====================
            tk.Label(adv_win, text="Tagname modèle", bg="white").pack(anchor="w", padx=10)
            model_tag_entry = ttk.Entry(adv_win)
            model_tag_entry.pack(fill="x", padx=10, pady=2)
        
            # ===================== BOUTONS HAUT =====================
            top_btn_frame = tk.Frame(adv_win, bg="white")
            top_btn_frame.pack(fill="x", padx=10, pady=5)
        
            tk.Button(
                top_btn_frame,
                text="Charger catégorie courante",
                command=lambda: load_model_current_category(),
                bg="#95a5a6", fg="white", relief="flat"
            ).pack(side="left", expand=True, fill="x", padx=2)
        
            tk.Button(
                top_btn_frame,
                text="Charger toutes les catégories",
                command=lambda: load_model_all_categories(),
                bg="#7f8c8d", fg="white", relief="flat"
            ).pack(side="left", expand=True, fill="x", padx=2)
        
            # ===================== CATEGORIE =====================
            tk.Label(adv_win, text="Catégorie", bg="white").pack(anchor="w", padx=10, pady=5)
            category_cb = ttk.Combobox(adv_win, values=list(categories.keys()), state="readonly")
            category_cb.pack(fill="x", padx=10)
            category_cb.current(0)
        
            # ===================== ZONE SCROLLABLE =====================
            fields_frame = tk.Frame(adv_win, bg="white")
            fields_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
            canvas = tk.Canvas(fields_frame, bg="white", highlightthickness=0)
            scrollbar = ttk.Scrollbar(fields_frame, orient="vertical", command=canvas.yview)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
        
            inner_frame = tk.Frame(canvas, bg="white")
            canvas.create_window((0, 0), window=inner_frame, anchor="nw")
            inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.configure(yscrollcommand=scrollbar.set)
        
            # ===================== DONNEES =====================
            temp_values = {}
            current_entries = {}
        
            # Injecter les valeurs sauvegardées précédemment
            if self.saved_advanced_params:
                temp_values.update(self.saved_advanced_params)
        
            # ===================== AFFICHAGE CATEGORIE =====================
            def show_category(cat_name, update_temp=True):
                nonlocal temp_values, current_entries
            
                if update_temp:
                    for col, entry in current_entries.items():
                        temp_values[col] = entry.get().strip()
            
                # Vider l'ancien affichage
                for w in inner_frame.winfo_children():
                    w.destroy()
                current_entries.clear()
            
                class_defaults = self.VAREXP_TEMPLATES.get(class_cb.get(), {})
                cols = categories[cat_name]
            
                for col in cols:
                    # Support des colonnes avec "..."
                    if col.endswith("..."):
                        prefix = col[:-3]
                        matching_cols = [h for h in self.headers if h.startswith(prefix)]
                    else:
                        matching_cols = [col]
            
                    for c in matching_cols:
                        tk.Label(inner_frame, text=c, bg="white").pack(anchor="w")
                        e = tk.Entry(inner_frame, relief="solid", bd=1)
                        e.pack(fill="x", pady=2)
            
                        # Afficher le commentaire si disponible
                        comment = CATEGORY_DESCRIPTIONS.get(c)
                        if comment:
                            tk.Label(inner_frame, text=comment, fg=self.COLORS["accent"], bg="white", font=("Segoe UI", 9, "italic")).pack(
                                anchor="w", padx=10, pady=(0, 5)
                            )
            
                        # Valeur initiale
                        if c in temp_values:
                            e.insert(0, temp_values[c])
                        elif c in class_defaults:
                            e.insert(0, class_defaults[c])
                        else:
                            e.insert(0, "")
                        
                        # ===== Mise en évidence initiale =====
                        template_val = class_defaults.get(c, "")
                        current_val = e.get()
                        if current_val != template_val:
                            e.config(bg="#fff6d5")  # jaune clair si modifié
                        else:
                            e.config(bg="white")     # blanc si identique au template
                        
                        # ===== Détection dynamique de modification =====
                        var = tk.StringVar(value=e.get())
                        e.config(textvariable=var)
                        
                        def on_entry_change(var, entry_widget=e, template_val=template_val):
                            val = var.get()
                            if val != template_val:
                                entry_widget.config(bg="#fff6d5")
                            else:
                                entry_widget.config(bg="white")
                        
                        var.trace_add("write", lambda *args, v=var: on_entry_change(v))
            
                        current_entries[c] = e
            
                canvas.yview_moveto(0)

            category_cb.bind("<<ComboboxSelected>>", lambda e: show_category(category_cb.get()))
        
            # ===================== RESET CATEGORIE COURANTE =====================
            def reset_current_category():
                class_defaults = self.VAREXP_TEMPLATES.get(class_cb.get(), {})
                current_category = category_cb.get()
        
                for col in categories[current_category]:
                    if col.endswith("..."):
                        prefix = col[:-3]
                        matching_cols = [h for h in self.headers if h.startswith(prefix)]
                    else:
                        matching_cols = [col]
        
                    for c in matching_cols:
                        if c in class_defaults:
                            temp_values[c] = class_defaults[c]
                        else:
                            temp_values.pop(c, None)
        
                        if c in current_entries:
                            current_entries[c].delete(0, tk.END)
                            if c in class_defaults:
                                current_entries[c].insert(0, class_defaults[c])
        
            # ===================== RESET TOUTES CATEGORIES =====================
            def reset_all_categories():
                class_defaults = self.VAREXP_TEMPLATES.get(class_cb.get(), {})
                temp_values.clear()
                temp_values.update(class_defaults)
                show_category(category_cb.get(), update_temp=False)
        
            # ===================== BOUTONS RESET =====================
            reset_btn_frame = tk.Frame(adv_win, bg="white")
            reset_btn_frame.pack(fill="x", padx=10, pady=5)
        
            tk.Button(
                reset_btn_frame,
                text="Réinitialiser la catégorie courante",
                command=reset_current_category,
                bg="#ffe4b5", relief="flat"
            ).pack(side="left", expand=True, fill="x", padx=2)
        
            tk.Button(
                reset_btn_frame,
                text="Réinitialiser toutes les catégories",
                command=reset_all_categories,
                bg="#ffcccc", relief="flat"
            ).pack(side="left", expand=True, fill="x", padx=2)
        
            # ===================== VALIDATION =====================
            def update_temp_values():
                for col, entry in current_entries.items():
                    temp_values[col] = entry.get().strip()
        
                advanced_params.clear()
                advanced_params.update(temp_values)
        
                if save_adv_var.get():
                    self.saved_advanced_params.clear()
                    self.saved_advanced_params.update(temp_values)
        
                adv_win.destroy()
        
            tk.Button(
                adv_win,
                text="Valider",
                command=update_temp_values,
                bg=self.COLORS["accent"], fg="white", relief="flat",
                font=("Segoe UI", 11, "bold")
            ).pack(pady=10)
        
            # ===================== CHARGEMENT MODELE =====================
            def extract_model_params(tagname):
                tag_col = self.headers.index("Tagname")
                model_row = None
        
                for row in self.data:
                    if len(row) > tag_col and row[tag_col] == str(tagname):
                        model_row = row
                        break
        
                if not model_row:
                    messagebox.showerror("Erreur", f"Tagname {tagname} introuvable")
                    return None
        
                model_values = {}
                for i, col in enumerate(self.headers):
                    if col in EXCLUDED_COLUMNS:
                        continue
                    if i < len(model_row) and model_row[i] != "":
                        model_values[col] = model_row[i]
        
                return model_values
        
            def load_model_current_category():
                tag = model_tag_entry.get().strip()
                if not tag.isdigit():
                    messagebox.showerror("Erreur", "Tagname invalide")
                    return
        
                model_values = extract_model_params(tag)
                if not model_values:
                    return
        
                current_category = category_cb.get()
                allowed_cols = categories[current_category]
        
                for col in allowed_cols:
                    if col in model_values:
                        temp_values[col] = model_values[col]
                        if col in current_entries:
                            current_entries[col].delete(0, tk.END)
                            current_entries[col].insert(0, model_values[col])
        
            def load_model_all_categories():
                tag = model_tag_entry.get().strip()
                if not tag.isdigit():
                    messagebox.showerror("Erreur", "Tagname invalide")
                    return
        
                model_values = extract_model_params(tag)
                if not model_values:
                    return
        
                temp_values.update(model_values)
                show_category(category_cb.get(), update_temp=False)
        
            # ===================== INIT =====================
            show_category(category_cb.get())

        tk.Button(win, text="Paramètres avancés", command=open_advanced, relief="flat", bg="#ecf0f1").pack(pady=10)

        
        # ---------- Validation ----------
        def validate():
            var_class = class_cb.get()
            var_name = name_entry.get().strip()
        
            # ----- Vérification du nom de variable -----
            if not var_name:
                win.attributes('-topmost', True)
                messagebox.showerror(
                    "Erreur", "Nom de variable obligatoire", parent=win
                )
                win.attributes('-topmost', False)
                return
            if not re.match(r'^[A-Za-z0-9_]+$', var_name):
                win.attributes('-topmost', True)
                messagebox.showerror(
                    "Erreur",
                    "Nom de variable invalide : seuls les lettres, chiffres et '_' sont autorisés",
                    parent=win
                )
                win.attributes('-topmost', False)
                return
        
            # ----- Construction du chemin n1..n12 -----
            raw_path = [e.get().strip() for e in path_entries if e.get().strip()]
        
            # ----- Vérification des éléments de chemin -----
            for elem in raw_path:
                if not re.match(r'^[A-Za-z0-9_]+$', elem):
                    win.attributes('-topmost', True)
                    messagebox.showerror(
                        "Erreur",
                        f"Élément de chemin invalide : '{elem}'. Seuls les lettres, chiffres et '_' sont autorisés",
                        parent=win
                    )
                    win.attributes('-topmost', False)
                    return
        
            if raw_path:
                path_elements = raw_path + [var_name]
            else:
                path_elements = [var_name]
        
            if len(path_elements) > 12:
                win.attributes('-topmost', True)
                messagebox.showerror(
                    "Erreur", "Chemin + nom de variable > 12 éléments", parent=win
                )
                win.attributes('-topmost', False)
                return
        
            # ===== Gestion des paramètres avancés =====
            if save_adv_var.get():
                adv_values = self.saved_advanced_params.copy()
                if advanced_params:
                    adv_values.update(advanced_params)
            else:
                adv_values = advanced_params.copy() if advanced_params else {}
        
            # ===== Vérification ServerListName / ClientListName =====
            server = adv_values.get("ServerListName", "").strip()
            client = adv_values.get("ClientListName", "").strip()
            if not server or not client:
                win.attributes('-topmost', True)
                confirm = messagebox.askyesno(
                    "Confirmation",
                    "ServerListName ou ClientListName est vide.\n"
                    "Êtes-vous sûr de vouloir créer cette variable ?",
                    parent=win
                )
                win.attributes('-topmost', False)
                if not confirm:
                    return
        
            # ===== Créer la variable =====
            self.create_variable(var_class, var_name, path_elements, adv_values)
        
            # ===== Réinitialiser les paramètres avancés si case décochée =====
            if not save_adv_var.get():
                advanced_params.clear()
                self.saved_advanced_params.clear()
                
        # Création du cadre pour aligner les boutons en bas de la fenêtre 'win'
        btn_frame = tk.Frame(win, bg="#ecf0f1", pady=10)
        btn_frame.pack(side='bottom', fill='x') # On le colle tout en bas

        # 1. Le bouton "Créer"
        tk.Button(btn_frame, 
                  text="Créer la variable", 
                  command=validate, 
                  bg=self.COLORS["success"], 
                  fg="white", 
                  relief="flat", 
                  font=("Segoe UI", 11, "bold")
        ).pack(side='left', padx=20, expand=True)

        # 2. Le bouton "Duplication de branche"
        tk.Button(btn_frame, 
                  text="Duplication de branche", 
                  command=self.open_duplicate_branch_window,
                  bg="#e67e22", 
                  fg="white", 
                  relief="flat",
                  font=("Segoe UI", 11, "bold")
        ).pack(side='right', padx=20, expand=True) 
            
    def create_variable(self, var_class, var_name, path_elements, adv_values=None):
        col_name = self.find_header("Nom")
        col_class = self.find_header("Class")
        col_tag = self.find_header("Tagname")
        
        new_row = [""] * len(self.headers)
        
        if col_name:
            new_row[self.headers.index(col_name)] = var_name
        if col_class:
            new_row[self.headers.index(col_class)] = var_class
        
        for i, elem in enumerate(path_elements[:11]):
            col = f"n{i+1}"
            if col in self.headers:
                new_row[self.headers.index(col)] = elem
        
        template = self.VAREXP_TEMPLATES.get(var_class, {})
        for col, val in template.items():
            if col in self.headers:
                new_row[self.headers.index(col)] = val
        
        if adv_values:
            for col, val in adv_values.items():
                if col in self.headers:
                    new_row[self.headers.index(col)] = val
        
        # =================================================================
        # C'EST ICI QU'ON MODIFIE POUR AVOIR LE BON TAG ID (> 65535)
        # =================================================================
        if col_tag:
            # On utilise la fonction qui recalcule le vrai MAX à chaque fois
            tag_id = self.get_next_tag_id()
            new_row[self.headers.index(col_tag)] = str(tag_id)
            
            # (Optionnel) On peut mettre à jour self.last_tagname si vous l'utilisez ailleurs
            self.last_tagname = tag_id 
        # =================================================================

        real_index = len(self.data)
        
        # Undo
        self.undo_stack.append([(real_index, None)])
        
        self.data.append(new_row)
        self.modified = True
        
        # Refresh filtre (recalcule les indices)
        self.apply_filter() 
        
        # Force le scroll tout en bas pour voir la nouvelle ligne créée
        self.scroll_bottom()

    # ================= COLONNES =================
    def select_columns(self):
        if not self.headers:
            return

        top = tk.Toplevel(self.root)
        top.title("Sélection des colonnes")
        top.geometry("500x800")
        top.configure(bg="white")

        col_vars = {}
        btn_frame = tk.Frame(top, bg="white")
        btn_frame.pack(fill="x", pady=5)

        def select_all():
            for var in col_vars.values():
                var.set(True)

        def deselect_all():
            for var in col_vars.values():
                var.set(False)

        def validate():
            self.visible_columns[:] = [c for c, v in col_vars.items() if v.get()]
            self.refresh_tree()
            for combobox in [self.column_filter1, self.column_filter2, self.column_filter3]:
                combobox['values'] = self.visible_columns
                if self.visible_columns:
                    combobox.current(0)
            top.destroy()

        tk.Button(btn_frame, text="Tout sélectionner", command=select_all, bg="#ecf0f1", relief="flat").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Tout désélectionner", command=deselect_all, bg="#ecf0f1", relief="flat").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Valider", command=validate, bg=self.COLORS["success"], fg="white", relief="flat").pack(side="right", padx=5)

        canvas = tk.Canvas(top, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(top, orient="vertical", command=canvas.yview)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)

        for col in self.headers:
            if col and "unnamed" not in col.lower():
                var = tk.BooleanVar(value=(col in self.visible_columns))
                tk.Checkbutton(frame, text=col, variable=var, bg="white").pack(anchor='w')
                col_vars[col] = var

    # ================= FILTRAGE (OPTIMISÉ) =================
    def reset_filters(self):
        # Vider les champs texte
        self.filter_entry1.delete(0, tk.END)
        self.filter_entry2.delete(0, tk.END)
        self.filter_entry3.delete(0, tk.END)
    
        # Réinitialiser les colonnes
        for cb in [self.column_filter1, self.column_filter2, self.column_filter3]:
            if self.visible_columns:
                cb.current(0)
    
        # Désactiver le filtrage
        self.filtered_indices = list(range(len(self.data)))
        self.refresh_tree()
    
    def apply_filter(self):
        # 1. Préparer les filtres actifs et pré-calculer les index de colonnes
        # Cela évite de faire .index() 50 000 fois dans la boucle
        raw_filters = [
            (self.column_filter1.get(), self.filter_entry1.get().lower().strip()),
            (self.column_filter2.get(), self.filter_entry2.get().lower().strip()),
            (self.column_filter3.get(), self.filter_entry3.get().lower().strip())
        ]
        
        active_filters = []
        for col_name, text in raw_filters:
            if col_name and text:
                try:
                    idx = self.headers.index(col_name)
                    active_filters.append((idx, text))
                except ValueError:
                    pass # La colonne n'existe pas

        mode = self.logic_mode.get()
        new_filtered_indices = []

        # 2. Si aucun filtre, on prend tout (plus rapide)
        if not active_filters:
            self.filtered_indices = list(range(len(self.data)))
            self.refresh_tree()
            return

        # 3. Boucle optimisée sur les index entiers
        for i, row in enumerate(self.data):
            matches = []
            for col_idx, text in active_filters:
                # Vérification directe par index
                if col_idx < len(row):
                    val = str(row[col_idx]).lower()
                    matches.append(text in val)
                else:
                    matches.append(False)
            
            if mode == "ET":
                if all(matches):
                    new_filtered_indices.append(i)
            else: # OU
                if any(matches):
                    new_filtered_indices.append(i)

        self.filtered_indices = new_filtered_indices
        self.refresh_tree()


if __name__ == "__main__":
    root = tk.Tk()
    DatEditor(root)
    root.mainloop()