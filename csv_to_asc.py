#!/usr/bin/env python3

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
        self.root.geometry("700x750")
        self.root.configure(bg='#f8e8f0')
        
        self.csv_file_path = tk.StringVar()
        self.gxo_file_path = tk.StringVar()
        self.deret_file_path = tk.StringVar()
        self.succ_gxo_file_path = tk.StringVar()
        self.succ_deret_file_path = tk.StringVar()
        self.aff_gxo_file_path = tk.StringVar()
        self.aff_deret_file_path = tk.StringVar()
        self.fra_gxo_file_path = tk.StringVar()
        self.fra_deret_file_path = tk.StringVar()
        
        self.document_type = tk.StringVar(value="TRF")
        self.column_mode = tk.StringVar(value="name")
        self.date_mode = tk.StringVar(value="auto")
        self.output_mode = tk.StringVar(value="simple")
        self.store_mode = tk.StringVar(value="file")
        self.sequence_counter = 70000
        
        # Variables globales pour le timestamp
        self.global_timestamp_base = None
        self.global_current_minutes = 0
        self.global_current_milliseconds = 894
        
        self.fra_stores = {599, 608, 609, 610, 611, 612, 613, 614, 615, 618, 619, 621, 625, 626, 628, 632, 633, 634, 635, 636, 638, 639, 640, 643, 648, 653, 655, 656, 657, 660, 663, 664, 665, 666, 667, 703}
        self.succ_exceptions = {600}  # Codes SUCC hors plage (> 399)
        
        tomorrow = datetime.now() + timedelta(days=1)
        self.selected_date_var = tk.StringVar(value=tomorrow.strftime("%d/%m/%Y"))
        self.manual_selected_date = tomorrow
        
        self.setup_ui()
        
    def setup_ui(self):
        self.bg_canvas = tk.Canvas(self.root, highlightthickness=0)
        self.bg_canvas.pack(fill=tk.BOTH, expand=True)
        
        self.main_canvas = tk.Canvas(self.bg_canvas, highlightthickness=0)
        main_scrollbar = ttk.Scrollbar(self.bg_canvas, orient="vertical", command=self.main_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.main_canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )
        
        self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=main_scrollbar.set)
        
        def create_gradient():
            self.bg_canvas.delete("gradient")
            width = self.bg_canvas.winfo_width()
            height = self.bg_canvas.winfo_height()
            if width > 1 and height > 1:
                r1, g1, b1 = 253, 242, 248
                r2, g2, b2 = 248, 215, 218
                
                for i in range(height):
                    ratio = i / height
                    r = int(r1 + (r2 - r1) * ratio)
                    g = int(g1 + (g2 - g1) * ratio)
                    b = int(b1 + (b2 - b1) * ratio)
                    color = f"#{r:02x}{g:02x}{b:02x}"
                    self.bg_canvas.create_line(0, i, width, i, fill=color, tags="gradient")
        
        def on_bg_configure(event=None):
            self.root.after_idle(create_gradient)
        
        self.bg_canvas.bind("<Configure>", on_bg_configure)
        self.root.after(1, create_gradient)
        
        main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        self.scrollable_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(9, weight=1)
        
        def configure_scroll_region(event=None):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
            canvas_width = self.bg_canvas.winfo_width()
            if canvas_width > 1:
                self.main_canvas.itemconfig(self.main_canvas.find_all()[0], width=canvas_width-20)
        
        self.bg_canvas.bind("<Configure>", lambda e: self.root.after_idle(configure_scroll_region))
        self.scrollable_frame.bind("<Configure>", configure_scroll_region)
        
        ttk.Label(main_frame, text="Convertisseur CSV vers ASC pour CEGID", 
                 font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        csv_frame = ttk.LabelFrame(main_frame, text="Fichier source", padding="10")
        csv_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        csv_frame.columnconfigure(1, weight=1)
        
        ttk.Label(csv_frame, text="Fichier CSV:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file_path).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(csv_frame, text="Parcourir", command=self.select_csv_file).grid(row=0, column=2, padx=(5, 0))
        
        self.asc_frame = ttk.LabelFrame(main_frame, text="Fichiers de sortie", padding="10")
        self.asc_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        self.asc_frame.columnconfigure(1, weight=1)
        
        self.create_output_widgets()
        
        mode_output_frame = ttk.LabelFrame(main_frame, text="Mode de sortie", padding="10")
        mode_output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Radiobutton(mode_output_frame, text="Simple", variable=self.output_mode, 
                       value="simple", command=self.on_output_mode_change).grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(mode_output_frame, text="Avancé", variable=self.output_mode, 
                       value="advanced", command=self.on_output_mode_change).grid(row=0, column=1, sticky=tk.W)
        
        store_mode_frame = ttk.LabelFrame(main_frame, text="Liste des magasins", padding="10")
        store_mode_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        store_mode_frame.columnconfigure(3, weight=1)
        
        ttk.Radiobutton(store_mode_frame, text="Par fichier (recommandé)", variable=self.store_mode,
                       value="file").grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(store_mode_frame, text="Hardcodée", variable=self.store_mode,
                       value="hardcoded").grid(row=0, column=1, sticky=tk.W, padx=(0, 15))
        ttk.Button(store_mode_frame, text="En savoir plus", command=self.show_store_mode_help,
                  style='Link.TButton').grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        
        mode_frame = ttk.LabelFrame(main_frame, text="Mode de sélection des colonnes", padding="10")
        mode_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        mode_frame.columnconfigure(3, weight=1)
        
        ttk.Radiobutton(mode_frame, text="Filtrage par nom", variable=self.column_mode, 
                       value="name").grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(mode_frame, text="Filtrage par ordre", variable=self.column_mode, 
                       value="order").grid(row=0, column=1, sticky=tk.W, padx=(0, 15))
        ttk.Button(mode_frame, text="En savoir plus", command=self.show_column_mode_help, 
                  style='Link.TButton').grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        
        date_frame = ttk.LabelFrame(main_frame, text="Mode de sélection de date", padding="10")
        date_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        date_frame.columnconfigure(3, weight=1)
        
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
        convert_button.grid(row=7, column=1, pady=10)
        
        ttk.Label(main_frame, text="Log de conversion:").grid(row=8, column=0, sticky=(tk.W, tk.N), pady=(5, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=10)
        scrollbar_log = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar_log.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_log.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        ttk.Label(main_frame, text="Fait par Clément :)", font=('Arial', 7), 
                 foreground='gray').grid(row=10, column=2, sticky=tk.E, pady=(2, 5))
        
        self.main_canvas.pack(side="left", fill="both", expand=True)
        main_scrollbar.pack(side="right", fill="y")
        
        self.bind_mousewheel(self.main_canvas)
        
        self.on_output_mode_change()
    
    def bind_mousewheel(self, canvas):
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def create_output_widgets(self):
        simple_configs = [
            ("GXO:", self.gxo_file_path, self.select_gxo_file),
            ("DERET:", self.deret_file_path, self.select_deret_file)
        ]
        
        advanced_configs = [
            ("SUCC GXO:", self.succ_gxo_file_path, self.select_succ_gxo_file),
            ("SUCC DERET:", self.succ_deret_file_path, self.select_succ_deret_file),
            ("AFF GXO:", self.aff_gxo_file_path, self.select_aff_gxo_file),
            ("AFF DERET:", self.aff_deret_file_path, self.select_aff_deret_file),
            ("FRA GXO:", self.fra_gxo_file_path, self.select_fra_gxo_file),
            ("FRA DERET:", self.fra_deret_file_path, self.select_fra_deret_file)
        ]
        
        self.simple_widgets = []
        self.advanced_widgets = []
        
        for widgets, configs in [(self.simple_widgets, simple_configs), (self.advanced_widgets, advanced_configs)]:
            for label, var, cmd in configs:
                label_widget = ttk.Label(self.asc_frame, text=label)
                entry_widget = ttk.Entry(self.asc_frame, textvariable=var)
                button_widget = ttk.Button(self.asc_frame, text="Parcourir", command=cmd)
                widgets.extend([label_widget, entry_widget, button_widget])
    
    def on_output_mode_change(self):
        for widget in self.simple_widgets + self.advanced_widgets:
            widget.grid_remove()
        
        widgets = self.simple_widgets if self.output_mode.get() == "simple" else self.advanced_widgets
        for i in range(0, len(widgets), 3):
            row = i // 3
            widgets[i].grid(row=row, column=0, sticky=tk.W, pady=5)
            widgets[i+1].grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
            widgets[i+2].grid(row=row, column=2, padx=(5, 0))
        
        self.root.after(10, self.update_scroll_region)
    
    def update_scroll_region(self):
        self.scrollable_frame.update_idletasks()
    
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
        
        first_day = calendar.monthrange(self.cal_year, self.cal_month)[0]
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

Avantages : Fonctionne même si l'ordre change, plus robuste et recommandé."""),
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

Important : L'ordre doit être exactement respecté, moins flexible que le filtrage par nom.""")
        ]
        
        for title, text in help_data:
            frame = ttk.LabelFrame(main_help_frame, text=title, padding="15")
            frame.pack(fill=tk.X, pady=(0, 15 if title == help_data[0][0] else 10))
            ttk.Label(frame, text=text, justify=tk.LEFT, wraplength=500).pack(anchor=tk.W)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def _on_mousewheel_col(e):
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel_col)
        help_window.protocol("WM_DELETE_WINDOW", lambda: [canvas.unbind_all("<MouseWheel>"), self.bind_mousewheel(self.main_canvas), help_window.destroy()])
        
    def show_store_mode_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Aide - Liste des magasins")
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
            ("Par fichier (recommandé)", """Le programme lit un fichier CSV pour déterminer le type de chaque magasin (SUCC, AFF ou FRA). Ce fichier doit être présent dans le même dossier que le script, et doit s'appeler "liste_magasins.csv"

Format attendu :
• Colonne A = code magasin (numérique à 3 chiffres)
• Colonne B = type (SUCC, AFF ou FRA)
Les autres colonnes sont ignorées.

Avantages : aucune modification du script nécessaire si un magasin change de statut ou si un nouveau magasin est ajouté."""),
            ("Hardcodée", """Le programme utilise les plages et listes de codes intégrées directement dans le script.

Règles appliquées :
• SUCC ≤ 399 (+ exception code 600)
• AFF 400–999
• FRA : 599, 608, 609, 610, 611, 612, 613, 614, 615, 618, 619, 621, 625, 626, 628, 632, 633, 634, 635, 636, 638, 639, 640, 643, 648, 653, 655, 656, 657, 660, 663, 664, 665, 666, 667, 703

Inconvénient : J'écris ce texte en Avril 2026, si la liste change d'ici là, alors cette liste sera obsolète et il faudra modifier le script pour qu'elle soit à jour. Je compte pas modifier le script à chaque mini changement (sauf contre virement de 20k pesos via PayPal). Pour cette raison, privilégiez la 1ère option.""")
        ]
        
        for title, text in help_data:
            frame = ttk.LabelFrame(main_help_frame, text=title, padding="15")
            frame.pack(fill=tk.X, pady=(0, 15 if title == help_data[0][0] else 10))
            ttk.Label(frame, text=text, justify=tk.LEFT, wraplength=500).pack(anchor=tk.W)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def _on_mousewheel_store(e):
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel_store)
        help_window.protocol("WM_DELETE_WINDOW", lambda: [canvas.unbind_all("<MouseWheel>"), self.bind_mousewheel(self.main_canvas), help_window.destroy()])

    def load_store_types_from_csv(self):
        """Charge le fichier liste_magasins.csv et retourne un dict {code: type} ou lève une ValueError."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(script_dir, "liste_magasins.csv")
        
        if not os.path.exists(csv_path):
            raise ValueError(
                "Fichier liste_magasins.csv introuvable.\n"
                f"Il doit être présent dans le même dossier que le script :\n{script_dir}\n\n"
                "Vérifiez que le fichier existe et que son nom est exactement \"liste_magasins.csv\"."
            )
        
        store_types = {}
        errors = []
        seen_codes = {}
        valid_types = {"SUCC", "AFF", "FRA"}
        
        for encoding in ['utf-8', 'cp1252', 'iso-8859-1', 'latin-1']:
            try:
                with open(csv_path, encoding=encoding, newline='') as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        import io
        import csv as csv_module
        reader = csv_module.reader(io.StringIO(content), delimiter=';')
        
        for line_num, row in enumerate(reader, start=1):
            if line_num == 1:
                # Ignorer l'en-tête si la première cellule n'est pas numérique
                try:
                    int(str(row[0]).strip())
                except (ValueError, IndexError):
                    continue
            
            if not row or all(cell.strip() == '' for cell in row):
                continue
            
            raw_code = str(row[0]).strip() if len(row) > 0 else ''
            raw_type = str(row[1]).strip().upper() if len(row) > 1 else ''
            
            # Vérification code numérique
            try:
                code = int(raw_code)
            except ValueError:
                errors.append(f"Ligne {line_num} : code non numérique \"{raw_code}\"")
                continue
            
            # Vérification doublon
            if code in seen_codes:
                errors.append(f"Ligne {line_num} : code {code} en doublon (déjà vu ligne {seen_codes[code]})")
                continue
            seen_codes[code] = line_num
            
            # Vérification statut manquant
            if not raw_type:
                errors.append(f"Ligne {line_num} : statut manquant pour le code {code}")
                continue
            
            # Vérification statut reconnu
            if raw_type not in valid_types:
                errors.append(f"Ligne {line_num} : statut \"{raw_type}\" non reconnu pour le code {code} (valeurs acceptées : SUCC, AFF, FRA)")
                continue
            
            store_types[code] = raw_type
        
        if errors:
            raise ValueError("Erreurs détectées dans liste_magasins.csv :\n\n" + "\n".join(f"• {e}" for e in errors))
        
        return store_types

    
        file_path = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.csv_file_path.set(file_path)
            base_name = os.path.splitext(file_path)[0]
            
            if self.output_mode.get() == "simple":
                self.gxo_file_path.set(f"{base_name}_GXO.asc")
                self.deret_file_path.set(f"{base_name}_DERET.asc")
            else:
                suffixes = [("_SUCC_GXO.asc", self.succ_gxo_file_path), ("_SUCC_DERET.asc", self.succ_deret_file_path),
                           ("_AFF_GXO.asc", self.aff_gxo_file_path), ("_AFF_DERET.asc", self.aff_deret_file_path),
                           ("_FRA_GXO.asc", self.fra_gxo_file_path), ("_FRA_DERET.asc", self.fra_deret_file_path)]
                for suffix, var in suffixes:
                    var.set(f"{base_name}{suffix}")
    
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
            if self.output_mode.get() == "advanced":
                suffixes = [("_SUCC_GXO.asc", self.succ_gxo_file_path), ("_SUCC_DERET.asc", self.succ_deret_file_path),
                           ("_AFF_GXO.asc", self.aff_gxo_file_path), ("_AFF_DERET.asc", self.aff_deret_file_path),
                           ("_FRA_GXO.asc", self.fra_gxo_file_path), ("_FRA_DERET.asc", self.fra_deret_file_path)]
                for suffix, var in suffixes:
                    var.set(f"{base_name}{suffix}")

    def select_gxo_file(self):
        self._select_file("Enregistrer le fichier GXO", self.gxo_file_path)
    def select_deret_file(self):
        self._select_file("Enregistrer le fichier DERET", self.deret_file_path)
    def select_succ_gxo_file(self):
        self._select_file("Enregistrer le fichier SUCC GXO", self.succ_gxo_file_path)
    def select_succ_deret_file(self):
        self._select_file("Enregistrer le fichier SUCC DERET", self.succ_deret_file_path)
    def select_aff_gxo_file(self):
        self._select_file("Enregistrer le fichier AFF GXO", self.aff_gxo_file_path)
    def select_aff_deret_file(self):
        self._select_file("Enregistrer le fichier AFF DERET", self.aff_deret_file_path)
    def select_fra_gxo_file(self):
        self._select_file("Enregistrer le fichier FRA GXO", self.fra_gxo_file_path)
    def select_fra_deret_file(self):
        self._select_file("Enregistrer le fichier FRA DERET", self.fra_deret_file_path)
    
    def _select_file(self, title, var):
        file_path = filedialog.asksaveasfilename(
            title=title, defaultextension=".asc",
            filetypes=[("Fichiers ASC", "*.asc"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            var.set(file_path)
    
    def get_store_type(self, store_code, store_map=None):
        try:
            code = int(store_code)
            if store_map is not None:
                return store_map.get(code, None)
            else:
                if code in self.fra_stores:
                    return "FRA"
                elif code <= 399 or code in self.succ_exceptions:
                    return "SUCC"
                elif 400 <= code <= 999:
                    return "AFF"
        except ValueError:
            pass
        return None
    
    def initialize_global_timestamp(self):
        """Initialise le timestamp global au début de la conversion"""
        self.global_timestamp_base = datetime.now()
        self.global_current_minutes = self.global_timestamp_base.minute
        self.global_current_milliseconds = 894
        self.log_message(f"Timestamp global initialisé : {self.global_timestamp_base.strftime('%H:%M')}.{self.global_current_milliseconds:03d}")
    
    def get_next_global_timestamp(self):
        """Génère le prochain timestamp global unique"""
        timestamp_str = f"{self.global_timestamp_base.year % 100:02d}{self.global_timestamp_base.day:02d}{self.global_timestamp_base.month:02d}:{self.global_timestamp_base.hour:02d}:{self.global_current_minutes:02d}.{self.global_current_milliseconds:03d}"
        
        # Incrémenter pour le prochain magasin
        self.global_current_milliseconds += 1
        if self.global_current_milliseconds > 999:
            self.global_current_milliseconds = 0
            self.global_current_minutes += 1
            if self.global_current_minutes > 59:
                self.global_current_minutes = 0
        
        return timestamp_str
            
    def get_next_sequence_number(self):
        sequence = self.sequence_counter
        self.sequence_counter += 1
        return f"{sequence}"
        
    def format_asc_header_line(self, store_code, sequence_number, date_str, timestamp_str):
        store_code_formatted = f"{store_code:03d}"
        line = f"E        {store_code_formatted}        {sequence_number}     {date_str}        {timestamp_str}                                                    EUR{date_str}"
        return line
        
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
        store_code_formatted = f"{store_code:03d}"
        
        if quantity < 100:
            quantity_formatted = f"   {quantity}"
        elif quantity < 1000:
            quantity_formatted = f"  {quantity}"
        else:
            quantity_formatted = f" {quantity}"
            
        return f"L{store_code_formatted}         {sequence_number}     {line_num_formatted}{date_str}                    {formatted_barcode}{quantity_formatted}"
        
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
                timestamp_str = self.get_next_global_timestamp()
                
                header_line = self.format_asc_header_line(store_code, sequence_number, date_str, timestamp_str)
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
                
                self.log_message(f"{file_type} - Magasin {store_code:03d}: {store_detail_count} articles traités (timestamp: {timestamp_str})")
        
        return total_stores, total_lines, total_pieces
        
    def convert_file(self):
        try:
            if self.output_mode.get() == "simple":
                required_files = [
                    (self.csv_file_path.get(), "Veuillez sélectionner un fichier CSV source"),
                    (self.gxo_file_path.get(), "Veuillez spécifier un fichier de sortie GXO"),
                    (self.deret_file_path.get(), "Veuillez spécifier un fichier de sortie DERET")
                ]
            else:
                required_files = [
                    (self.csv_file_path.get(), "Veuillez sélectionner un fichier CSV source"),
                    (self.succ_gxo_file_path.get(), "Veuillez spécifier un fichier de sortie SUCC GXO"),
                    (self.succ_deret_file_path.get(), "Veuillez spécifier un fichier de sortie SUCC DERET"),
                    (self.aff_gxo_file_path.get(), "Veuillez spécifier un fichier de sortie AFF GXO"),
                    (self.aff_deret_file_path.get(), "Veuillez spécifier un fichier de sortie AFF DERET"),
                    (self.fra_gxo_file_path.get(), "Veuillez spécifier un fichier de sortie FRA GXO"),
                    (self.fra_deret_file_path.get(), "Veuillez spécifier un fichier de sortie FRA DERET")
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
            
            # Initialiser le timestamp global au début de la conversion
            self.initialize_global_timestamp()
            
            if self.output_mode.get() == "simple":
                self.log_message("Début de la conversion avec séparation GXO/DERET...")
            else:
                self.log_message("Début de la conversion avec séparation SUCC/AFF/FRA et GXO/DERET...")
                
            self.log_message(f"Date utilisée: {selected_date.strftime('%d/%m/%Y')} (Mode: {self.date_mode.get()})")
            self.log_message(f"Lecture du fichier CSV: {self.csv_file_path.get()}")
            
            # Chargement de la liste des magasins si mode avancé
            store_map = None
            if self.output_mode.get() == "advanced":
                if self.store_mode.get() == "file":
                    self.log_message("Chargement de liste_magasins.csv...")
                    store_map = self.load_store_types_from_csv()
                    self.log_message(f"liste_magasins.csv chargé : {len(store_map)} magasins")
                else:
                    self.log_message("Mode hardcodé : utilisation des plages et listes intégrées")
            
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
            
            if self.output_mode.get() == "simple":
                df_gxo = df[df[column_mapping['best']] == 'GXO'].copy()
                df_deret = df[df[column_mapping['best']] != 'GXO'].copy()
                
                self.log_message(f"Lignes GXO: {len(df_gxo)}")
                self.log_message(f"Lignes DERET: {len(df_deret)}")
                
                results = {}
                # Ordre de génération pour le mode simple : GXO puis DERET
                file_configs = [
                    (df_gxo, self.gxo_file_path.get(), "GXO"),
                    (df_deret, self.deret_file_path.get(), "DERET")
                ]
                
                for data, path, file_type in file_configs:
                    if len(data) > 0:
                        self.log_message(f"Génération du fichier {file_type}...")
                        results[file_type] = self.generate_asc_file(data, path, file_type, column_mapping, selected_date)
                        self.log_message(f"Fichier {file_type} généré: {results[file_type][0]} magasins, {results[file_type][1]} lignes, {results[file_type][2]} pièces")
                    else:
                        self.log_message(f"Aucune donnée {file_type} trouvée, fichier {file_type} non généré")
                        results[file_type] = (0, 0, 0)
                
                success_message = (f"Conversion terminée avec succès!\n\n"
                                  f"Date utilisée: {selected_date.strftime('%d/%m/%Y')}\n"
                                  f"Fichier GXO: {results['GXO'][0]} magasins, {results['GXO'][1]} lignes, {results['GXO'][2]} pièces\n"
                                  f"Fichier DERET: {results['DERET'][0]} magasins, {results['DERET'][1]} lignes, {results['DERET'][2]} pièces")
            else:
                df['store_type'] = df[column_mapping['store']].apply(lambda x: self.get_store_type(x, store_map))
                df_valid = df[df['store_type'].notna()].copy()
                excluded_count = len(df) - len(df_valid)
                
                if excluded_count > 0:
                    self.log_message(f"Magasins exclus (codes >800 ou non reconnus): {excluded_count} lignes")
                
                self.log_message(f"Lignes avec magasins valides: {len(df_valid)}")
                
                df_succ_gxo = df_valid[(df_valid['store_type'] == 'SUCC') & (df_valid[column_mapping['best']] == 'GXO')].copy()
                df_succ_deret = df_valid[(df_valid['store_type'] == 'SUCC') & (df_valid[column_mapping['best']] != 'GXO')].copy()
                df_aff_gxo = df_valid[(df_valid['store_type'] == 'AFF') & (df_valid[column_mapping['best']] == 'GXO')].copy()
                df_aff_deret = df_valid[(df_valid['store_type'] == 'AFF') & (df_valid[column_mapping['best']] != 'GXO')].copy()
                df_fra_gxo = df_valid[(df_valid['store_type'] == 'FRA') & (df_valid[column_mapping['best']] == 'GXO')].copy()
                df_fra_deret = df_valid[(df_valid['store_type'] == 'FRA') & (df_valid[column_mapping['best']] != 'GXO')].copy()
                
                for name, df_subset in [("SUCC GXO", df_succ_gxo), ("SUCC DERET", df_succ_deret), 
                                       ("AFF GXO", df_aff_gxo), ("AFF DERET", df_aff_deret),
                                       ("FRA GXO", df_fra_gxo), ("FRA DERET", df_fra_deret)]:
                    self.log_message(f"Lignes {name}: {len(df_subset)}")
                
                results = {}
                # Ordre de génération pour le mode avancé : SUCC GXO → SUCC DERET → AFF GXO → AFF DERET → FRA GXO → FRA DERET
                file_configs = [
                    (df_succ_gxo, self.succ_gxo_file_path.get(), "SUCC GXO"),
                    (df_succ_deret, self.succ_deret_file_path.get(), "SUCC DERET"),
                    (df_aff_gxo, self.aff_gxo_file_path.get(), "AFF GXO"),
                    (df_aff_deret, self.aff_deret_file_path.get(), "AFF DERET"),
                    (df_fra_gxo, self.fra_gxo_file_path.get(), "FRA GXO"),
                    (df_fra_deret, self.fra_deret_file_path.get(), "FRA DERET")
                ]
                
                for data, path, file_type in file_configs:
                    if len(data) > 0:
                        self.log_message(f"Génération du fichier {file_type}...")
                        results[file_type] = self.generate_asc_file(data, path, file_type, column_mapping, selected_date)
                        self.log_message(f"Fichier {file_type} généré: {results[file_type][0]} magasins, {results[file_type][1]} lignes, {results[file_type][2]} pièces")
                    else:
                        self.log_message(f"Aucune donnée {file_type} trouvée, fichier {file_type} non généré")
                        results[file_type] = (0, 0, 0)
                
                success_message = (f"Conversion terminée avec succès!\n\n"
                                  f"Date utilisée: {selected_date.strftime('%d/%m/%Y')}\n"
                                  f"Fichier SUCC GXO: {results['SUCC GXO'][0]} magasins, {results['SUCC GXO'][1]} lignes, {results['SUCC GXO'][2]} pièces\n"
                                  f"Fichier SUCC DERET: {results['SUCC DERET'][0]} magasins, {results['SUCC DERET'][1]} lignes, {results['SUCC DERET'][2]} pièces\n"
                                  f"Fichier AFF GXO: {results['AFF GXO'][0]} magasins, {results['AFF GXO'][1]} lignes, {results['AFF GXO'][2]} pièces\n"
                                  f"Fichier AFF DERET: {results['AFF DERET'][0]} magasins, {results['AFF DERET'][1]} lignes, {results['AFF DERET'][2]} pièces\n"
                                  f"Fichier FRA GXO: {results['FRA GXO'][0]} magasins, {results['FRA GXO'][1]} lignes, {results['FRA GXO'][2]} pièces\n"
                                  f"Fichier FRA DERET: {results['FRA DERET'][0]} magasins, {results['FRA DERET'][1]} lignes, {results['FRA DERET'][2]} pièces")
            
            self.log_message("Conversion terminée avec succès!")
            self.log_message(f"Timestamp final : {self.global_timestamp_base.strftime('%H:%M')}.{self.global_current_milliseconds:03d}")
            messagebox.showinfo("Succès", success_message)
                              
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