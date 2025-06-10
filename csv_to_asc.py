#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime, date, timedelta
import calendar
import os

class CSVToASCConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur CSV vers ASC - CEGID")
        self.root.geometry("700x600")
        self.root.configure(bg='#f8e8f0')
        
        self.csv_file_path = tk.StringVar()
        self.gxo_file_path = tk.StringVar()
        self.deret_file_path = tk.StringVar()
        self.document_type = tk.StringVar(value="TRF")
        self.column_mode = tk.StringVar(value="name")
        self.date_mode = tk.StringVar(value="auto")
        self.sequence_counter = 163406
        
        tomorrow = datetime.now() + timedelta(days=1)
        self.selected_date_var = tk.StringVar(value=tomorrow.strftime("%d/%m/%Y"))
        self.manual_selected_date = tomorrow
        
        self.setup_ui()
        
    def setup_ui(self):
        canvas = tk.Canvas(self.root, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        def create_gradient(canvas, color1, color2, width, height):
            r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
            r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)
            for i in range(height):
                ratio = i / height
                color = f"#{int(r1+(r2-r1)*ratio):02x}{int(g1+(g2-g1)*ratio):02x}{int(b1+(b2-b1)*ratio):02x}"
                canvas.create_line(0, i, width, i, fill=color)
        
        def update_gradient(event=None):
            canvas.delete("all")
            width, height = canvas.winfo_width(), canvas.winfo_height()
            if width > 1 and height > 1:
                create_gradient(canvas, "#fdf2f8", "#f8d7da", width, height)
        
        canvas.bind("<Configure>", update_gradient)
        self.root.after(1, update_gradient)
        
        main_frame = ttk.Frame(canvas, padding="10")
        main_frame.place(x=0, y=0, relwidth=1, relheight=1)
        
        for i in [0, 1]: 
            (self.root if i == 0 else main_frame).columnconfigure(i, weight=1)
            (self.root if i == 0 else main_frame).rowconfigure(i, weight=1 if i == 0 else 0)
        canvas.columnconfigure(0, weight=1)
        canvas.rowconfigure(0, weight=1)
        main_frame.rowconfigure(7, weight=1)
        
        ttk.Label(main_frame, text="Convertisseur CSV vers ASC pour CEGID", 
                 font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        csv_frame = ttk.LabelFrame(main_frame, text="Fichier source", padding="10")
        csv_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        csv_frame.columnconfigure(1, weight=1)
        
        ttk.Label(csv_frame, text="Fichier CSV:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(csv_frame, text="Parcourir", command=self.select_csv_file).grid(row=0, column=2, padx=(5, 0))
        
        asc_frame = ttk.LabelFrame(main_frame, text="Fichiers de sortie", padding="10")
        asc_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        asc_frame.columnconfigure(1, weight=1)
        
        for i, (label, var, cmd) in enumerate([
            ("Fichier ASC GXO:", self.gxo_file_path, self.select_gxo_file),
            ("Fichier ASC DERET:", self.deret_file_path, self.select_deret_file)
        ]):
            ttk.Label(asc_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
            ttk.Entry(asc_frame, textvariable=var, width=50).grid(row=i, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
            ttk.Button(asc_frame, text="Parcourir", command=cmd).grid(row=i, column=2, padx=(5, 0))
        
        mode_frame = ttk.LabelFrame(main_frame, text="Mode de sélection des colonnes", padding="10")
        mode_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        mode_frame.columnconfigure(2, weight=1)
        
        ttk.Radiobutton(mode_frame, text="Filtrage par nom", variable=self.column_mode, 
                       value="name").grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(mode_frame, text="Filtrage par ordre", variable=self.column_mode, 
                       value="order").grid(row=0, column=1, sticky=tk.W, padx=(0, 15))
        ttk.Button(mode_frame, text="En savoir plus", command=self.show_column_mode_help, 
                  style='Link.TButton').grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        
        date_frame = ttk.LabelFrame(main_frame, text="Mode de sélection de date", padding="10")
        date_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        date_frame.columnconfigure(2, weight=1)
        
        tomorrow_str = (datetime.now() + timedelta(days=1)).strftime("%d/%m")
        ttk.Radiobutton(date_frame, text=f"Automatique ({tomorrow_str})", 
                       variable=self.date_mode, value="auto", 
                       command=self.on_date_mode_change).grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(date_frame, text="Manuel", variable=self.date_mode, 
                       value="manual", command=self.on_date_mode_change).grid(row=0, column=1, sticky=tk.W, padx=(0, 15))
        
        self.calendar_button = ttk.Button(date_frame, text="Choisir une date", 
                                         command=self.open_calendar, state='disabled')
        self.calendar_button.grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        
        self.selected_date_label = ttk.Label(date_frame, textvariable=self.selected_date_var, foreground='gray')
        self.selected_date_label.grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        
        convert_button = ttk.Button(main_frame, text="Convertir", command=self.convert_file, style='Accent.TButton')
        convert_button.configure(width=20)
        convert_button.grid(row=5, column=1, pady=10)
        
        ttk.Label(main_frame, text="Log de conversion:").grid(row=6, column=0, sticky=(tk.W, tk.N), pady=(5, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=10, width=80)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        ttk.Label(main_frame, text="Fait par Clément :)", font=('Arial', 7), 
                 foreground='gray').grid(row=8, column=2, sticky=tk.E, pady=(10, 0))
    
    def on_date_mode_change(self):
        if self.date_mode.get() == "manual":
            self.calendar_button.config(state='normal')
            self.selected_date_label.config(foreground='black')
        else:
            self.calendar_button.config(state='disabled')
            self.selected_date_label.config(foreground='gray')
            tomorrow = datetime.now() + timedelta(days=1)
            self.selected_date_var.set(tomorrow.strftime("%d/%m/%Y"))
    
    def open_calendar(self):
        cal_window = tk.Toplevel(self.root)
        cal_window.title("Sélectionner une date")
        cal_window.geometry("350x260")
        cal_window.configure(bg='white')
        cal_window.transient(self.root)
        cal_window.grab_set()
        cal_window.resizable(False, False)
        cal_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        current_date = self.manual_selected_date
        self.cal_year, self.cal_month, self.cal_selected_day = current_date.year, current_date.month, current_date.day
        
        main_cal_frame = ttk.Frame(cal_window)
        main_cal_frame.pack(fill=tk.BOTH, expand=True)
        
        header_frame = ttk.Frame(main_cal_frame)
        header_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Button(header_frame, text="◀", width=3, 
                  command=lambda: self.change_month(-1, cal_window)).pack(side=tk.LEFT)
        self.month_year_label = ttk.Label(header_frame, font=('Arial', 12, 'bold'))
        self.month_year_label.pack(side=tk.LEFT, expand=True)
        ttk.Button(header_frame, text="▶", width=3, 
                  command=lambda: self.change_month(1, cal_window)).pack(side=tk.RIGHT)
        
        self.calendar_frame = ttk.Frame(main_cal_frame)
        self.calendar_frame.pack(padx=10, pady=(0, 5))
        
        self.generate_calendar(cal_window)
    
    def change_month(self, direction, cal_window):
        self.cal_month += direction
        if self.cal_month > 12:
            self.cal_month, self.cal_year = 1, self.cal_year + 1
        elif self.cal_month < 1:
            self.cal_month, self.cal_year = 12, self.cal_year - 1
        self.generate_calendar(cal_window)
    
    def generate_calendar(self, cal_window):
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        months = ['', 'janvier', 'février', 'mars', 'avril', 'mai', 'juin',
                 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
        self.month_year_label.config(text=f"{months[self.cal_month]} {self.cal_year}")
        
        for i, day in enumerate(['Lu', 'Ma', 'Me', 'Je', 'Ve', 'Sa', 'Di']):
            ttk.Label(self.calendar_frame, text=day, font=('Arial', 9, 'bold')).grid(
                row=0, column=i, padx=1, pady=1, ipadx=6, ipady=2)
        
        first_day = (calendar.monthrange(self.cal_year, self.cal_month)[0] + 1) % 7
        days_in_month = calendar.monthrange(self.cal_year, self.cal_month)[1]
        tomorrow = date.today() + timedelta(days=1)
        
        row, col = 1, first_day
        for day in range(1, days_in_month + 1):
            day_date = date(self.cal_year, self.cal_month, day)
            
            if day_date < tomorrow:
                btn = tk.Button(self.calendar_frame, text=str(day), width=3, height=1,
                               state='disabled', bg='#f0f0f0', fg='gray')
            elif day_date == tomorrow:
                btn = tk.Button(self.calendar_frame, text=str(day), width=3, height=1,
                               bg='#e3f2fd', fg='#1976d2', font=('Arial', 9, 'bold'),
                               command=lambda d=day: self.select_day(d, cal_window))
            else:
                btn = tk.Button(self.calendar_frame, text=str(day), width=3, height=1,
                               bg='white', fg='black',
                               command=lambda d=day: self.select_day(d, cal_window))
            
            btn.grid(row=row, column=col, padx=1, pady=1, ipadx=1, ipady=1)
            col += 1
            if col > 6:
                col, row = 0, row + 1
    
    def select_day(self, day, cal_window):
        selected_date = datetime(self.cal_year, self.cal_month, day)
        tomorrow = date.today() + timedelta(days=1)
        
        if selected_date.date() < tomorrow:
            messagebox.showwarning("Date invalide", "Vous ne pouvez pas sélectionner une date antérieure à demain.")
            return
        
        self.manual_selected_date = selected_date
        self.selected_date_var.set(selected_date.strftime("%d/%m/%Y"))
        cal_window.destroy()
    
    def get_selected_date(self):
        return datetime.now() + timedelta(days=1) if self.date_mode.get() == "auto" else self.manual_selected_date
        
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def show_column_mode_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Aide - Modes de sélection des colonnes")
        help_window.geometry("580x580")
        help_window.configure(bg='#f8e8f0')
        help_window.transient(self.root)
        help_window.grab_set()
        help_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        canvas = tk.Canvas(help_window, bg='#f8e8f0')
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        main_help_frame = ttk.Frame(scrollable_frame, padding="20")
        main_help_frame.pack(fill=tk.BOTH, expand=True)
        
        help_data = [
            ("Filtrage par nom (recommandé)", """Le programme recherche les colonnes par leur nom exact dans le fichier CSV.

Colonnes recherchées :
• Magasin : "Recipient store" ou "Etab."
• Code-barres : "Code-barres article"
• Quantité : "Quantité saisie transfert"
• Type : "BEST"

✅ Avantages : Fonctionne même si l'ordre change, plus robuste et recommandé."""),
            ("Filtrage par ordre", """Le programme utilise la position des colonnes dans le fichier CSV.

Ordre attendu (basé sur le format REA.csv) :
1. Recipient store (magasin)
2. Code article
3. Code-barres article
4. Libellé article
5. Quantité saisie transfert
6. Stock net
7. Stock initial
8. Qté vendue
9. Stock mini
10. Stock maxi
11. CONDITIONNEMENT
12. Stock dispo.
13. BEST

⚠️ Important : L'ordre doit être exactement respecté, moins flexible que le filtrage par nom.""")
        ]
        
        for title, text in help_data:
            frame = ttk.LabelFrame(main_help_frame, text=title, padding="15")
            frame.pack(fill=tk.X, pady=(0, 15 if title == help_data[0][0] else 10))
            ttk.Label(frame, text=text, justify=tk.LEFT).pack(anchor=tk.W)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
    def select_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.csv_file_path.set(file_path)
            base_name = os.path.splitext(file_path)[0]
            self.gxo_file_path.set(f"{base_name}_GXO.asc")
            self.deret_file_path.set(f"{base_name}_DERET.asc")
            
    def select_gxo_file(self):
        self._select_file("Enregistrer le fichier ASC GXO", self.gxo_file_path)
            
    def select_deret_file(self):
        self._select_file("Enregistrer le fichier ASC DERET", self.deret_file_path)
    
    def _select_file(self, title, var):
        file_path = filedialog.asksaveasfilename(
            title=title, defaultextension=".asc",
            filetypes=[("Fichiers ASC", "*.asc"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            var.set(file_path)
            
    def get_next_sequence_number(self):
        sequence = self.sequence_counter
        self.sequence_counter += 1
        return f"{sequence:06d}"
        
    def format_asc_header_line(self, store_code, sequence_number, date_str):
        doc_type = "TRF_PGI" if self.document_type.get() == "TRF" else "CBR_CLI"
        line = f"E        {store_code:03d}          {sequence_number}  {date_str}        {doc_type} {date_str}                                                    EUR{date_str}"
        return line.ljust(191)
        
    def format_barcode(self, barcode):
        clean_barcode = ''.join(filter(str.isdigit, str(barcode).strip()))
        if len(clean_barcode) >= 11:
            clean_barcode = "0137" + clean_barcode[-11:]
        else:
            clean_barcode = "0137" + clean_barcode.zfill(11)
        return clean_barcode[:15].zfill(15)
        
    def format_asc_detail_line(self, store_code, sequence_number, line_number, date_str, barcode, quantity):
        formatted_barcode = self.format_barcode(barcode)
        line_num_formatted = f"{line_number:04d}"
        return f"L{store_code:03d}     INTINT{sequence_number}  {line_num_formatted}{date_str}                    {formatted_barcode}  {quantity:>3d}            "
        
    def find_store_column(self, df):
        if self.column_mode.get() == "name":
            for col in ["Recipient store", "Etab."]:
                if col in df.columns:
                    self.log_message(f"Mode nom - Colonne magasin trouvée: '{col}'")
                    return col
            raise ValueError(f"Mode nom - Aucune colonne magasin trouvée. Colonnes disponibles: {list(df.columns)}. "
                            f"Colonnes acceptées: ['Recipient store', 'Etab.']")
        else:
            if len(df.columns) >= 1:
                store_column = df.columns[0]
                self.log_message(f"Mode ordre - Colonne magasin (position 0): '{store_column}'")
                return store_column
            raise ValueError("Mode ordre - Le fichier CSV doit avoir au moins 1 colonne")
    
    def get_required_columns(self, df):
        if self.column_mode.get() == "name":
            store_column = self.find_store_column(df)
            column_mapping = {
                'store': store_column,
                'barcode': "Code-barres article",
                'quantity': "Quantité saisie transfert", 
                'best': "BEST"
            }
            missing_columns = [col for key, col in column_mapping.items() if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Mode nom - Colonnes manquantes: {missing_columns}")
        else:
            required_positions = {'store': 0, 'barcode': 2, 'quantity': 4, 'best': 12}
            max_position = max(required_positions.values())
            if len(df.columns) <= max_position:
                raise ValueError(f"Mode ordre - Le fichier CSV doit avoir au moins {max_position + 1} colonnes. "
                               f"Actuellement: {len(df.columns)} colonnes")
            
            column_mapping = {key: df.columns[pos] for key, pos in required_positions.items()}
            self.log_message("Mode ordre - Colonnes utilisées:")
            for key, pos in required_positions.items():
                self.log_message(f"  {key.title()} (pos {pos}): '{column_mapping[key]}'")
        
        return column_mapping
        
    def generate_asc_file(self, df_filtered, output_path, file_type, column_mapping, selected_date):
        self.sequence_counter = 163406
        
        with open(output_path, 'w', encoding='utf-8', newline='') as asc_file:
            total_lines = total_stores = total_pieces = 0
            date_str = selected_date.strftime("%d/%m/%y")
            
            for store_code_str, store_data in df_filtered.groupby(column_mapping['store']):
                try:
                    store_code = int(store_code_str)
                except ValueError:
                    self.log_message(f"Code magasin invalide '{store_code_str}', ignoré ({file_type})")
                    continue
                    
                sequence_number = self.get_next_sequence_number()
                header_line = self.format_asc_header_line(store_code, sequence_number, date_str)
                asc_file.write(header_line + '\r\n')
                total_lines += 1
                total_stores += 1
                
                line_number = 1
                store_detail_count = 0
                
                for _, row in store_data.iterrows():
                    try:
                        barcode = str(row[column_mapping['barcode']]).strip()
                        quantity_str = str(row[column_mapping['quantity']]).strip()
                        
                        if not barcode or barcode in ['nan', '']:
                            continue
                            
                        try:
                            quantity = int(float(quantity_str))
                            if quantity <= 0:
                                continue
                        except (ValueError, TypeError):
                            continue
                        
                        detail_line = self.format_asc_detail_line(store_code, sequence_number, line_number, date_str, barcode, quantity)
                        asc_file.write(detail_line + '\r\n')
                        
                        line_number += 1
                        store_detail_count += 1
                        total_lines += 1
                        total_pieces += quantity
                        
                    except Exception:
                        continue
                
                self.log_message(f"{file_type} - Magasin {store_code:03d}: {store_detail_count} articles traités")
        
        return total_stores, total_lines, total_pieces
        
    def convert_file(self):
        try:
            required_files = [
                (self.csv_file_path.get(), "Veuillez sélectionner un fichier CSV source"),
                (self.gxo_file_path.get(), "Veuillez spécifier un fichier de sortie GXO"),
                (self.deret_file_path.get(), "Veuillez spécifier un fichier de sortie DERET")
            ]
            
            for file_path, error_msg in required_files:
                if not file_path:
                    messagebox.showerror("Erreur", error_msg)
                    return
            
            selected_date = self.get_selected_date()
            tomorrow = date.today() + timedelta(days=1)
            if self.date_mode.get() == "manual" and selected_date.date() < tomorrow:
                messagebox.showerror("Erreur", "La date sélectionnée ne peut pas être antérieure à demain")
                return
                
            self.log_text.delete(1.0, tk.END)
            self.log_message("Début de la conversion avec séparation GXO/DERET...")
            self.log_message(f"Type de document: {self.document_type.get()}")
            self.log_message(f"Date utilisée: {selected_date.strftime('%d/%m/%Y')} (Mode: {self.date_mode.get()})")
            self.log_message(f"Lecture du fichier CSV: {self.csv_file_path.get()}")
            
            df = None
            for encoding in ['utf-8', 'cp1252', 'iso-8859-1', 'latin-1']:
                try:
                    self.log_message(f"Tentative de lecture avec l'encodage: {encoding}")
                    df = pd.read_csv(self.csv_file_path.get(), sep=';', dtype=str, encoding=encoding)
                    self.log_message(f"Succès avec l'encodage: {encoding}")
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    if encoding == 'latin-1':
                        raise e
            
            if df is None:
                raise ValueError("Impossible de lire le fichier CSV avec les encodages supportés")
            
            df.columns = df.columns.str.strip()
            self.log_message(f"Fichier CSV lu avec succès: {len(df)} lignes")
            self.log_message(f"Colonnes trouvées: {list(df.columns)}")
            self.log_message(f"Mode de sélection: {self.column_mode.get()}")
            
            column_mapping = self.get_required_columns(df)
            
            df = df.dropna(subset=[column_mapping['quantity']])
            df = df[df[column_mapping['quantity']].astype(str).str.strip().isin(['']) == False]
            df = df[df[column_mapping['quantity']].astype(str).str.strip() != '0']
            
            self.log_message(f"Après filtrage: {len(df)} lignes avec quantités valides")
            
            df[column_mapping['best']] = df[column_mapping['best']].fillna('DERET').astype(str).str.strip()
            df.loc[df[column_mapping['best']].isin(['', 'nan']), column_mapping['best']] = 'DERET'
            
            df_gxo = df[df[column_mapping['best']] == 'GXO'].copy()
            df_deret = df[df[column_mapping['best']] != 'GXO'].copy()
            
            self.log_message(f"Lignes GXO: {len(df_gxo)}")
            self.log_message(f"Lignes DERET: {len(df_deret)}")
            
            results = {}
            for data, path, file_type in [(df_gxo, self.gxo_file_path.get(), "GXO"), 
                                         (df_deret, self.deret_file_path.get(), "DERET")]:
                if len(data) > 0:
                    self.log_message(f"Génération du fichier {file_type}...")
                    results[file_type] = self.generate_asc_file(data, path, file_type, column_mapping, selected_date)
                    self.log_message(f"Fichier {file_type} généré: {results[file_type][0]} magasins, {results[file_type][1]} lignes, {results[file_type][2]} pièces")
                else:
                    self.log_message(f"Aucune donnée {file_type} trouvée, fichier {file_type} non généré")
                    results[file_type] = (0, 0, 0)
            
            self.log_message("Conversion terminée avec succès!")
            
            messagebox.showinfo("Succès", 
                              f"Conversion terminée avec succès!\n\n"
                              f"Date utilisée: {selected_date.strftime('%d/%m/%Y')}\n"
                              f"Fichier GXO: {results['GXO'][0]} magasins, {results['GXO'][1]} lignes, {results['GXO'][2]} pièces\n"
                              f"Fichier DERET: {results['DERET'][0]} magasins, {results['DERET'][1]} lignes, {results['DERET'][2]} pièces")
                              
        except Exception as e:
            error_msg = f"Erreur lors de la conversion: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Erreur", error_msg)

def main():
    root = tk.Tk()
    app = CSVToASCConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()