#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os

class CSVToASCConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur CSV vers ASC - CEGID")
        self.root.geometry("700x550")
        self.root.configure(bg='#f8e8f0')
        
        self.csv_file_path = tk.StringVar()
        self.gxo_file_path = tk.StringVar()
        self.deret_file_path = tk.StringVar()
        self.document_type = tk.StringVar(value="TRF")
        
        self.sequence_counter = 163406
        
        self.setup_ui()
        
    def setup_ui(self):
        # Création d'un canvas pour le dégradé
        canvas = tk.Canvas(self.root, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Création du dégradé rose
        def create_gradient(canvas, color1, color2, width, height):
            r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
            r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)
            
            for i in range(height):
                ratio = i / height
                r = int(r1 + (r2 - r1) * ratio)
                g = int(g1 + (g2 - g1) * ratio)
                b = int(b1 + (b2 - b1) * ratio)
                color = f"#{r:02x}{g:02x}{b:02x}"
                canvas.create_line(0, i, width, i, fill=color)
        
        # Mise à jour du dégradé lors du redimensionnement
        def update_gradient(event=None):
            canvas.delete("all")
            width = canvas.winfo_width()
            height = canvas.winfo_height()
            if width > 1 and height > 1:
                create_gradient(canvas, "#fdf2f8", "#f8d7da", width, height)
        
        canvas.bind("<Configure>", update_gradient)
        self.root.after(1, update_gradient)
        
        main_frame = ttk.Frame(canvas, padding="10")
        main_frame.place(x=0, y=0, relwidth=1, relheight=1)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        canvas.columnconfigure(0, weight=1)
        canvas.rowconfigure(0, weight=1)
        
        title_label = ttk.Label(main_frame, text="Convertisseur CSV vers ASC pour CEGID", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Section fichier CSV source
        csv_frame = ttk.LabelFrame(main_frame, text="Fichier source", padding="10")
        csv_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        csv_frame.columnconfigure(1, weight=1)
        
        ttk.Label(csv_frame, text="Fichier CSV:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(csv_frame, text="Parcourir", command=self.select_csv_file).grid(row=0, column=2, padx=(5, 0))
        
        # Section fichiers ASC de sortie
        asc_frame = ttk.LabelFrame(main_frame, text="Fichiers de sortie", padding="10")
        asc_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        asc_frame.columnconfigure(1, weight=1)
        
        ttk.Label(asc_frame, text="Fichier ASC GXO:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(asc_frame, textvariable=self.gxo_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(asc_frame, text="Parcourir", command=self.select_gxo_file).grid(row=0, column=2, padx=(5, 0))
        
        ttk.Label(asc_frame, text="Fichier ASC DERET:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(asc_frame, textvariable=self.deret_file_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(asc_frame, text="Parcourir", command=self.select_deret_file).grid(row=1, column=2, padx=(5, 0))
        
        convert_button = ttk.Button(main_frame, text="Convertir", command=self.convert_file, 
                                   style='Accent.TButton')
        convert_button.configure(width=20)
        convert_button.grid(row=3, column=1, pady=10)
        
        ttk.Label(main_frame, text="Log de conversion:").grid(row=4, column=0, sticky=(tk.W, tk.N), pady=(5, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=12, width=80)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        main_frame.rowconfigure(5, weight=1)
        
        credit_label = ttk.Label(main_frame, text="Fait par Clément :)", 
                                font=('Arial', 7), foreground='gray')
        credit_label.grid(row=6, column=2, sticky=tk.E, pady=(10, 0))
        
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
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
        file_path = filedialog.asksaveasfilename(
            title="Enregistrer le fichier ASC GXO",
            defaultextension=".asc",
            filetypes=[("Fichiers ASC", "*.asc"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.gxo_file_path.set(file_path)
            
    def select_deret_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Enregistrer le fichier ASC DERET",
            defaultextension=".asc",
            filetypes=[("Fichiers ASC", "*.asc"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.deret_file_path.set(file_path)
            
    def get_next_sequence_number(self):
        sequence = self.sequence_counter
        self.sequence_counter += 1
        return f"{sequence:06d}"
        
    def format_asc_header_line(self, store_code, sequence_number, date_str):
        doc_type = f"TRF_PGI" if self.document_type.get() == "TRF" else "CBR_CLI"
        
        line = f"E        {store_code:03d}          {sequence_number}  {date_str}        {doc_type} {date_str}                                                    EUR{date_str}"
        
        line = line.ljust(191)
        return line
        
    def format_barcode(self, barcode):
        clean_barcode = str(barcode).strip()
        
        clean_barcode = ''.join(filter(str.isdigit, clean_barcode))
        
        if len(clean_barcode) >= 11:
            clean_barcode = "0137" + clean_barcode[-11:]
        elif len(clean_barcode) < 11:
            clean_barcode = "0137" + clean_barcode.zfill(11)
        
        return clean_barcode[:15].zfill(15)
        
    def format_asc_detail_line(self, store_code, sequence_number, line_number, date_str, barcode, quantity):
        
        formatted_barcode = self.format_barcode(barcode)
        
        line_num_formatted = f"{line_number:04d}"
        
        line = f"L{store_code:03d}     INTINT{sequence_number}  {line_num_formatted}{date_str}                    {formatted_barcode}  {quantity:>3d}            "
        
        return line
        
    def generate_asc_file(self, df_filtered, output_path, file_type):
        self.sequence_counter = 163406
        
        with open(output_path, 'w', encoding='utf-8', newline='') as asc_file:
            total_lines = 0
            total_stores = 0
            total_pieces = 0
            
            current_date = datetime.now()
            date_str = current_date.strftime("%d/%m/%y")
            
            grouped = df_filtered.groupby('Recipient store')
            
            for store_code_str, store_data in grouped:
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
                        barcode = str(row['Code-barres article']).strip()
                        quantity_str = str(row['Quantité saisie transfert']).strip()
                        
                        if not barcode or barcode == 'nan' or barcode == '':
                            continue
                            
                        try:
                            quantity = int(float(quantity_str))
                            if quantity <= 0:
                                continue
                        except (ValueError, TypeError):
                            continue
                        
                        detail_line = self.format_asc_detail_line(
                            store_code, sequence_number, line_number, 
                            date_str, barcode, quantity
                        )
                        asc_file.write(detail_line + '\r\n')
                        
                        line_number += 1
                        store_detail_count += 1
                        total_lines += 1
                        total_pieces += quantity
                        
                    except Exception as e:
                        continue
                
                self.log_message(f"{file_type} - Magasin {store_code:03d}: {store_detail_count} articles traités")
        
        return total_stores, total_lines, total_pieces
        
    def convert_file(self):
        try:
            if not self.csv_file_path.get():
                messagebox.showerror("Erreur", "Veuillez sélectionner un fichier CSV source")
                return
                
            if not self.gxo_file_path.get():
                messagebox.showerror("Erreur", "Veuillez spécifier un fichier de sortie GXO")
                return
                
            if not self.deret_file_path.get():
                messagebox.showerror("Erreur", "Veuillez spécifier un fichier de sortie DERET")
                return
                
            self.log_text.delete(1.0, tk.END)
            self.log_message("Début de la conversion avec séparation GXO/DERET...")
            self.log_message(f"Type de document: {self.document_type.get()}")
            
            self.log_message(f"Lecture du fichier CSV: {self.csv_file_path.get()}")
            
            encodings_to_try = ['utf-8', 'cp1252', 'iso-8859-1', 'latin-1']
            df = None
            
            for encoding in encodings_to_try:
                try:
                    self.log_message(f"Tentative de lecture avec l'encodage: {encoding}")
                    df = pd.read_csv(self.csv_file_path.get(), sep=';', dtype=str, encoding=encoding)
                    self.log_message(f"Succès avec l'encodage: {encoding}")
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    if encoding == encodings_to_try[-1]:
                        raise e
                    continue
            
            if df is None:
                raise ValueError("Impossible de lire le fichier CSV avec les encodages supportés")
            
            df.columns = df.columns.str.strip()
            
            self.log_message(f"Fichier CSV lu avec succès: {len(df)} lignes")
            self.log_message(f"Colonnes trouvées: {list(df.columns)}")
            
            required_columns = ["Recipient store", "Code-barres article", "Quantité saisie transfert", "BEST"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Colonnes manquantes dans le CSV: {missing_columns}")
            
            df = df.dropna(subset=['Quantité saisie transfert'])
            df = df[df['Quantité saisie transfert'].astype(str).str.strip() != '']
            df = df[df['Quantité saisie transfert'].astype(str).str.strip() != '0']
            
            self.log_message(f"Après filtrage: {len(df)} lignes avec quantités valides")
            
            df['BEST'] = df['BEST'].fillna('DERET')
            df['BEST'] = df['BEST'].astype(str).str.strip()
            df.loc[df['BEST'] == '', 'BEST'] = 'DERET'
            df.loc[df['BEST'] == 'nan', 'BEST'] = 'DERET'
            
            df_gxo = df[df['BEST'] == 'GXO'].copy()
            df_deret = df[df['BEST'] != 'GXO'].copy()
            
            self.log_message(f"Lignes GXO: {len(df_gxo)}")
            self.log_message(f"Lignes DERET: {len(df_deret)}")
            
            gxo_stats = (0, 0, 0)
            deret_stats = (0, 0, 0)
            
            if len(df_gxo) > 0:
                self.log_message("Génération du fichier GXO...")
                gxo_stats = self.generate_asc_file(df_gxo, self.gxo_file_path.get(), "GXO")
                self.log_message(f"Fichier GXO généré: {gxo_stats[0]} magasins, {gxo_stats[1]} lignes, {gxo_stats[2]} pièces")
            else:
                self.log_message("Aucune donnée GXO trouvée, fichier GXO non généré")
            
            if len(df_deret) > 0:
                self.log_message("Génération du fichier DERET...")
                deret_stats = self.generate_asc_file(df_deret, self.deret_file_path.get(), "DERET")
                self.log_message(f"Fichier DERET généré: {deret_stats[0]} magasins, {deret_stats[1]} lignes, {deret_stats[2]} pièces")
            else:
                self.log_message("Aucune donnée DERET trouvée, fichier DERET non généré")
            
            self.log_message("Conversion terminée avec succès!")
            
            messagebox.showinfo("Succès", 
                              f"Conversion terminée avec succès!\n\n"
                              f"Fichier GXO: {gxo_stats[0]} magasins, {gxo_stats[1]} lignes, {gxo_stats[2]} pièces\n"
                              f"Fichier DERET: {deret_stats[0]} magasins, {deret_stats[1]} lignes, {deret_stats[2]} pièces")
                              
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