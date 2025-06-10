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
        self.root.geometry("700x450")
        
        self.csv_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.document_type = tk.StringVar(value="TRF")
        
        self.sequence_counter = 163406
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        title_label = ttk.Label(main_frame, text="Convertisseur CSV vers ASC pour CEGID", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        ttk.Label(main_frame, text="Fichier CSV source:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.csv_file_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(main_frame, text="Parcourir", command=self.select_csv_file).grid(row=1, column=2, padx=(5, 0))
        
        ttk.Label(main_frame, text="Fichier ASC de sortie:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_file_path, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(main_frame, text="Parcourir", command=self.select_output_file).grid(row=2, column=2, padx=(5, 0))
        
        convert_button = ttk.Button(main_frame, text="Convertir", command=self.convert_file, 
                                   style='Accent.TButton')
        convert_button.grid(row=3, column=1, pady=15)
        
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
            self.output_file_path.set(f"{base_name}.asc")
            
    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Enregistrer le fichier ASC",
            defaultextension=".asc",
            filetypes=[("Fichiers ASC", "*.asc"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.output_file_path.set(file_path)
            
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
        
    def convert_file(self):
        try:
            if not self.csv_file_path.get():
                messagebox.showerror("Erreur", "Veuillez sélectionner un fichier CSV source")
                return
                
            if not self.output_file_path.get():
                messagebox.showerror("Erreur", "Veuillez spécifier un fichier de sortie")
                return
                
            self.sequence_counter = 163406
                
            self.log_text.delete(1.0, tk.END)
            self.log_message("Début de la conversion avec le format CEGID...")
            self.log_message(f"Type de document: {self.document_type.get()}")
            self.log_message(f"Séquence de départ: {self.sequence_counter}")
            
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
            
            required_columns = ["Recipient store", "Code-barres article", "Quantité saisie transfert"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Colonnes manquantes dans le CSV: {missing_columns}")
            
            df = df.dropna(subset=['Quantité saisie transfert'])
            df = df[df['Quantité saisie transfert'].astype(str).str.strip() != '']
            df = df[df['Quantité saisie transfert'].astype(str).str.strip() != '0']
            
            self.log_message(f"Après filtrage: {len(df)} lignes avec quantités valides")
            
            current_date = datetime.now()
            date_str = current_date.strftime("%d/%m/%y")
            self.log_message(f"Date utilisée: {date_str}")
            
            grouped = df.groupby('Recipient store')
            self.log_message(f"Données groupées par {len(grouped)} magasins")
            
            self.log_message("Génération du fichier ASC avec format CEGID exact...")
            
            with open(self.output_file_path.get(), 'w', encoding='utf-8', newline='') as asc_file:
                total_lines = 0
                total_stores = 0
                total_pieces = 0
                
                for store_code_str, store_data in grouped:
                    try:
                        store_code = int(store_code_str)
                    except ValueError:
                        self.log_message(f"Code magasin invalide '{store_code_str}', ignoré")
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
                                self.log_message(f"Code-barre manquant, ligne ignorée")
                                continue
                                
                            try:
                                quantity = int(float(quantity_str))
                                if quantity <= 0:
                                    self.log_message(f"Quantité invalide ({quantity}), ligne ignorée")
                                    continue
                            except (ValueError, TypeError):
                                self.log_message(f"Quantité non numérique ({quantity_str}), ligne ignorée")
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
                            self.log_message(f"Erreur sur une ligne du magasin {store_code}: {e}")
                            continue
                    
                    self.log_message(f"Magasin {store_code:03d}: {store_detail_count} articles traités (séquence: {sequence_number})")
            
            self.log_message(f"Conversion terminée avec succès!")
            self.log_message(f"Fichier généré: {self.output_file_path.get()}")
            self.log_message(f"Total: {total_stores} magasins, {total_lines} lignes générées")
            
            self.validate_generated_file()
            
            messagebox.showinfo("Succès", 
                              f"Conversion terminée avec succès!\n"
                              f"Fichier généré: {self.output_file_path.get()}\n"
                              f"Total: {total_stores} magasins, {total_lines} lignes, {total_pieces} pièces")
                              
        except Exception as e:
            error_msg = f"Erreur lors de la conversion: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Erreur", error_msg)
            
    def validate_generated_file(self):
        try:
            self.log_message("Validation du fichier généré selon les standards CEGID...")
            
            with open(self.output_file_path.get(), 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            header_count = 0
            detail_count = 0
            errors = []
            warnings = []
            sequence_numbers = set()
            
            for i, line in enumerate(lines, 1):
                line = line.rstrip('\r\n')
                
                if line.startswith('E'):
                    header_count += 1
                    if len(line) != 191:
                        errors.append(f"Ligne {i}: En-tête longueur incorrecte ({len(line)} au lieu de 191 caractères)")
                        
                elif line.startswith('L'):
                    detail_count += 1
                    
                    try:
                        seq_start = line.find('INTINT') + 6
                        seq_end = seq_start + 6
                        sequence_num = line[seq_start:seq_end]
                        sequence_numbers.add(sequence_num)
                    except:
                        pass
                    
                    if len(line) < 70:
                        errors.append(f"Ligne {i}: Ligne de détail trop courte ({len(line)} caractères)")
                        continue
                    
                    try:
                        barcode_start = line.find('                    ') + 20
                        if barcode_start > 19:
                            barcode_section = line[barcode_start:barcode_start+15].strip()
                            barcode_digits = ''.join(filter(str.isdigit, barcode_section))
                            
                            if len(barcode_digits) != 15:
                                warnings.append(f"Ligne {i}: Code-barre non standard ({len(barcode_digits)} chiffres au lieu de 15) - '{barcode_section}'")
                            elif not barcode_digits.startswith('0137'):
                                warnings.append(f"Ligne {i}: Code-barre ne commence pas par '0137' - '{barcode_digits}'")
                    except Exception as e:
                        errors.append(f"Ligne {i}: Code-barre illisible - {e}")
                    
                    try:
                        quantity_section = line[70:].strip()
                        if not quantity_section:
                            errors.append(f"Ligne {i}: Quantité manquante (erreur QTE_T1)")
                        else:
                            quantity_digits = ''.join(filter(str.isdigit, quantity_section.split()[0] if quantity_section.split() else ''))
                            if not quantity_digits:
                                errors.append(f"Ligne {i}: Quantité non numérique (erreur QTE_T1)")
                    except:
                        errors.append(f"Ligne {i}: Erreur dans la section quantité (QTE_T1)")
            
            self.log_message(f"Validation: {header_count} en-têtes, {detail_count} lignes de détail")
            self.log_message(f"{len(sequence_numbers)} numéros de séquence différents générés")
            
            if errors:
                self.log_message("Erreurs critiques détectées:")
                for error in errors[:5]:
                    self.log_message(f"  {error}")
                if len(errors) > 5:
                    self.log_message(f"  ... et {len(errors) - 5} autres erreurs critiques")
                    
            if warnings:
                self.log_message("Avertissements:")
                for warning in warnings[:3]:
                    self.log_message(f"  {warning}")
                if len(warnings) > 3:
                    self.log_message(f"  ... et {len(warnings) - 3} autres avertissements")
            
            if not errors:
                self.log_message("Validation réussie: format compatible CEGID")
                
        except Exception as e:
            self.log_message(f"Erreur lors de la validation: {e}")

def main():
    root = tk.Tk()
    app = CSVToASCConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()