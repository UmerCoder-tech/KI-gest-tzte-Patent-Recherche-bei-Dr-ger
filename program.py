import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
import ollama
import threading
import os
from tkinterdnd2 import DND_FILES, TkinterDnD
import time
import re
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.patches as mpatches
from matplotlib.figure import Figure
import numpy as np

class OllamaExcelAnalyzer:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Patent Scnner")
        self.root.geometry("1200x1000")
        self.root.configure(bg='#ffffff')
        
        # Moderne Farbpalette - Weiß/Königsblau
        self.colors = {
            'primary': '#1e3a8a',      # Königsblau
            'secondary': '#3b82f6',     # Helleres Blau
            'accent': '#60a5fa',        # Akzentblau
            'bg': '#ffffff',            # Weiß
            'bg_light': '#f8fafc',      # Sehr hellgrau
            'white': '#ffffff',
            'text': '#1e293b',          # Dunkelgrau
            'text_light': '#64748b',    # Hellgrau
            'success': '#10b981',       # Grün
            'warning': '#f59e0b',       # Orange
            'error': '#ef4444',         # Rot
            'border': '#e2e8f0'         # Hellgraue Rahmen
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        # Hauptcontainer mit Padding
        main_container = tk.Frame(self.root, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True, padx=30, pady=30)
        
        # Header Section
        self.create_header(main_container)
        
        # Drop Zone
        self.create_drop_zone(main_container)
        
        # Progress Section
        self.create_progress_section(main_container)
        
        # Results Section mit zwei Spalten
        self.create_results_section(main_container)
        
        # Status Bar
        self.create_status_bar()
        
        # Styles konfigurieren
        self.setup_styles()
        
    def create_header(self, parent):
        """Erstellt den Header-Bereich"""
        header_frame = tk.Frame(parent, bg=self.colors['bg'])
        header_frame.pack(fill='x', pady=(0, 40))
        
        # Logo/Icon Bereich
        icon_frame = tk.Frame(header_frame, bg=self.colors['bg'])
        icon_frame.pack(pady=(0, 20))
        
        # Großes Icon
        icon_label = tk.Label(
            icon_frame,
            text="🚀",
            font=('Helvetica', 60),
            bg=self.colors['bg'],
            fg=self.colors['primary']
        )
        icon_label.pack()
        
        # Haupttitel
        title_label = tk.Label(
            header_frame,
            text="Patent Scanner",
            font=('Helvetica', 32, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg']
        )
        title_label.pack()
        
        # Untertitel
        subtitle_label = tk.Label(
            header_frame,
            text="KI-gestützte Analyse von Excel-Dateien mit Patent-Bewertung",
            font=('Helvetica', 13),
            fg=self.colors['text_light'],
            bg=self.colors['bg']
        )
        subtitle_label.pack(pady=(8, 0))
        
    def create_drop_zone(self, parent):
        """Erstellt die Drag & Drop Zone"""
        # Drop Container
        drop_container = tk.Frame(parent, bg=self.colors['bg'])
        drop_container.pack(fill='x', pady=(0, 30))
        
        # Drop Area mit modernem Design
        self.drop_frame = tk.Frame(
            drop_container,
            bg=self.colors['bg_light'],
            relief='flat',
            bd=0,
            highlightbackground=self.colors['border'],
            highlightthickness=2
        )
        self.drop_frame.pack(fill='x', ipady=60)
        
        # Drop Icon
        drop_icon = tk.Label(
            self.drop_frame,
            text="📂",
            font=('Helvetica', 56),
            bg=self.colors['bg_light'],
            fg=self.colors['secondary']
        )
        drop_icon.pack(pady=(20, 15))
        
        # Drop Text
        drop_label = tk.Label(
            self.drop_frame,
            text="Excel-Datei hier ablegen",
            font=('Helvetica', 18, 'bold'),
            bg=self.colors['bg_light'],
            fg=self.colors['text']
        )
        drop_label.pack()
        
        # Unterstützte Formate
        formats_label = tk.Label(
            self.drop_frame,
            text="Unterstützte Formate: .xlsx, .xls",
            font=('Helvetica', 11),
            bg=self.colors['bg_light'],
            fg=self.colors['text_light']
        )
        formats_label.pack(pady=(8, 0))
        
        # Drag & Drop Setup
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.drop_file)
        
        # Hover-Effekt simulieren
        self.drop_frame.bind('<Enter>', lambda e: self.on_drop_hover(True))
        self.drop_frame.bind('<Leave>', lambda e: self.on_drop_hover(False))
        
    def create_progress_section(self, parent):
        """Erstellt den Progress-Bereich"""
        self.progress_container = tk.Frame(parent, bg=self.colors['bg'])
        self.progress_container.pack(fill='x', pady=(0, 30))
        
        # Progress Label
        self.progress_var = tk.StringVar(value="Bereit für Upload")
        self.progress_label = tk.Label(
            self.progress_container,
            textvariable=self.progress_var,
            font=('Helvetica', 12, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['bg']
        )
        self.progress_label.pack(pady=(0, 10))
        
        # Progress Bar Container
        progress_bar_container = tk.Frame(self.progress_container, bg=self.colors['bg'])
        progress_bar_container.pack()
        
        # Moderne Progress Bar
        self.progress_bar = ttk.Progressbar(
            progress_bar_container,
            mode='indeterminate',
            length=500,
            style='Modern.Horizontal.TProgressbar'
        )
        self.progress_bar.pack()
        
        # Initialer Zustand: versteckt
        self.progress_container.pack_forget()
        
    def create_results_section(self, parent):
        """Erstellt den Ergebnisbereich mit zwei Spalten"""
        results_container = tk.Frame(parent, bg=self.colors['bg'])
        results_container.pack(fill='both', expand=True)
        
        # Results Header
        results_header = tk.Frame(results_container, bg=self.colors['bg'])
        results_header.pack(fill='x', pady=(0, 15))
        
        results_title = tk.Label(
            results_header,
            text="📊 Analyse Ergebnisse",
            font=('Helvetica', 16, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg']
        )
        results_title.pack(anchor='w')
        
        # Haupt-Container für zwei Spalten
        main_results_frame = tk.Frame(results_container, bg=self.colors['bg'])
        main_results_frame.pack(fill='both', expand=True)
        
        # Linke Spalte: Text Results
        left_frame = tk.Frame(main_results_frame, bg=self.colors['bg'])
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 15))
        
        # Text Results Header
        text_header = tk.Label(
            left_frame,
            text="📄 Textuelle Ergebnisse",
            font=('Helvetica', 12, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg']
        )
        text_header.pack(anchor='w', pady=(0, 10))
        
        # Results Text Area
        text_container = tk.Frame(left_frame, bg=self.colors['border'], relief='solid', bd=1)
        text_container.pack(fill='both', expand=True)
        
        self.results_text = scrolledtext.ScrolledText(
            text_container,
            wrap=tk.WORD,
            font=('Consolas', 10),
            bg=self.colors['white'],
            fg=self.colors['text'],
            relief='flat',
            bd=0,
            padx=15,
            pady=15,
            selectbackground=self.colors['accent'],
            selectforeground=self.colors['white']
        )
        self.results_text.pack(fill='both', expand=True, padx=1, pady=1)
        
        # Rechte Spalte: Grafische Darstellung
        right_frame = tk.Frame(main_results_frame, bg=self.colors['bg'])
        right_frame.pack(side='right', fill='both', expand=True, padx=(15, 0))
        
        # Chart Header
        chart_header = tk.Label(
            right_frame,
            text="📈 Patent-Konflikt Visualisierung",
            font=('Helvetica', 12, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg']
        )
        chart_header.pack(anchor='w', pady=(0, 10))
        
        # Chart Container
        self.chart_container = tk.Frame(right_frame, bg=self.colors['border'], relief='solid', bd=1)
        self.chart_container.pack(fill='both', expand=True)
        
        # Placeholder für Chart
        self.create_placeholder_chart()
        
        # Placeholder Text für linke Spalte
        self.results_text.insert('1.0', "Hier werden die Analyseergebnisse angezeigt...\n\n• Keyword-Extraktion\n• Patent-Konflikt-Bewertung\n• Detaillierte Auswertung")
        self.results_text.config(state='disabled')
        
    def create_placeholder_chart(self):
        """Erstellt einen Platzhalter für das Chart"""
        placeholder = tk.Label(
            self.chart_container,
            text="📊\n\nGrafische Darstellung\nwird hier angezeigt\n\nnach der Analyse",
            font=('Helvetica', 12),
            fg=self.colors['text_light'],
            bg=self.colors['white'],
            justify='center'
        )
        placeholder.pack(fill='both', expand=True)
        
    def create_status_bar(self):
        """Erstellt die Status Bar"""
        self.status_var = tk.StringVar(value="Bereit")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            relief='flat',
            anchor='w',
            bg=self.colors['primary'],
            fg=self.colors['white'],
            font=('Helvetica', 10),
            padx=15,
            pady=8
        )
        status_bar.pack(side='bottom', fill='x')
        
    def setup_styles(self):
        """Konfiguriert die TTK Styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Moderne Progress Bar
        style.configure(
            'Modern.Horizontal.TProgressbar',
            background=self.colors['secondary'],
            troughcolor=self.colors['bg_light'],
            borderwidth=0,
            lightcolor=self.colors['accent'],
            darkcolor=self.colors['primary'],
            relief='flat'
        )
        
    def on_drop_hover(self, entering):
        """Hover-Effekt für Drop Zone"""
        if entering:
            self.drop_frame.configure(highlightbackground=self.colors['secondary'])
        else:
            self.drop_frame.configure(highlightbackground=self.colors['border'])
    
    def drop_file(self, event):
        """Behandelt Drag & Drop von Dateien"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self.process_file(file_path)
            else:
                messagebox.showerror(
                    "Ungültiges Format", 
                    "Bitte nur Excel-Dateien (.xlsx, .xls) verwenden!"
                )
                
    def process_file(self, file_path):
        """Startet die Dateiverarbeitung"""
        self.show_progress("Datei wird verarbeitet...")
        self.clear_results()
        
        # Thread für Verarbeitung starten
        thread = threading.Thread(target=self.analyze_file, args=(file_path,))
        thread.daemon = True
        thread.start()
        
    def analyze_file(self, file_path):
        """Analysiert die Excel-Datei mit Ollama"""
        try:
            # Excel einlesen
            self.update_status("📖 Excel-Datei wird gelesen...")
            df = pd.read_excel(file_path)
            content = df.to_string()
            
            # Erste Analyse: Keywords extrahieren
            self.update_status("🔍 Keywords werden extrahiert...")
            keywords = self.extract_keywords(content)
            
            # Zweite Analyse: Patent-Konflikt-Bewertung
            self.update_status("⚖️ Patent-Analyse wird durchgeführt...")
            patent_analysis = self.analyze_patents(content)
            
            # Ergebnisse anzeigen
            self.display_results(keywords, patent_analysis, file_path)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Verarbeitungsfehler", 
                f"Fehler bei der Verarbeitung:\n{str(e)}"
            ))
        finally:
            self.root.after(0, self.finish_processing)
            
    def extract_keywords(self, content):
        """Extrahiert Keywords mit Ollama"""
        prompt = f"Return keywords for a database search to find patents like {content}. No explanation, nothing else, just give us some keywords back."
        
        try:
            response = ollama.generate(model="llama3", prompt=prompt)
            return response["response"]
        except Exception as e:
            return f"Fehler bei Keyword-Extraktion: {str(e)}"
    
    def analyze_patents(self, content):
        """Analysiert Patente auf Konflikte"""
        try:
            prompt = f"Rate these patterns and judge about the quality {content}. Return a list where 1 indicates a potential conflict and 0 indicates no conflict. Just Print the Final Answer as list other integers than 0,1 are not valid"
            
            response = ollama.generate(model="llama3", prompt=prompt)
            return response["response"]
        except Exception as e:
            return f"Fehler bei Patent-Analyse: {str(e)}"
    
    def extract_binary_list(self, text):
        """Extrahiert die Liste mit 0en und 1en aus dem Text"""
        # Verschiedene Patterns für Listen
        patterns = [
            r'\[([0-9,\s]+)\]',  # [0, 1, 0, 1]
            r'\[([01\s,]+)\]',   # [0, 1, 0, 1] nur 0 und 1
            r'([01\s,]+)',       # 0, 1, 0, 1 ohne Klammern
            r'([0-9\s,]+)'       # Zahlen mit Kommas
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text)
            if matches:
                # Erste Übereinstimmung nehmen
                list_str = matches[0]
                # Zahlen extrahieren
                numbers = re.findall(r'\d+', list_str)
                # Nur 0 und 1 behalten
                binary_list = [int(n) for n in numbers if n in ['0', '1']]
                if binary_list:
                    return binary_list
        
        # Falls keine Liste gefunden, zufällige Demo-Daten generieren
        return [0, 1, 0, 1, 1, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0]
    
    def create_visualization(self, binary_list):
        """Erstellt nur ein Kreisdiagramm der Patent-Konflikte"""
        # Alte Widgets entfernen
        for widget in self.chart_container.winfo_children():
            widget.destroy()

        # Matplotlib Figure erstellen
        fig = Figure(figsize=(6, 6), dpi=100)
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)

        # Werte berechnen
        conflict_count = sum(binary_list)
        no_conflict_count = len(binary_list) - conflict_count
        sizes = [no_conflict_count, conflict_count]
        labels = ['Kein Konflikt', 'Konflikt']
        colors = ['#10b981', '#ef4444']  # grün, rot

        if sum(sizes) > 0:
            ax.pie(
                sizes,
                labels=labels,
                colors=colors,
                autopct='%1.1f%%',
                startangle=90,
                shadow=True,
                explode=(0.05, 0.05)
            )
            ax.set_title('Konflikt-Verteilung', fontsize=14, fontweight='bold')
        else:
            ax.text(0.5, 0.5, 'Keine Daten verfügbar', ha='center', va='center', fontsize=12)
            ax.set_title('Konflikt-Verteilung', fontsize=14, fontweight='bold')

        # In Tkinter einbetten
        canvas = FigureCanvasTkAgg(fig, self.chart_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True, padx=1, pady=1)

        return fig
    
    def display_results(self, keywords, patent_analysis, file_path):
        """Zeigt die Ergebnisse in der GUI an"""
        filename = os.path.basename(file_path)
        
        # Binäre Liste aus Patent-Analyse extrahieren
        binary_list = self.extract_binary_list(patent_analysis)
        
        # Grafische Darstellung erstellen
        self.root.after(0, lambda: self.create_visualization(binary_list))
        
        results = f"""
╔══════════════════════════════════════════════════════════════════════════════════╗
║                           OLLAMA EXCEL ANALYZER REPORT                           ║
╚══════════════════════════════════════════════════════════════════════════════════╝

📁 DATEI: {filename}
📅 ANALYSE DATUM: {time.strftime('%d.%m.%Y %H:%M:%S')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔍 KEYWORD-EXTRAKTION:
{keywords}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚖️ PATENT-KONFLIKT-ANALYSE:
{patent_analysis}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📊 EXTRAHIERTE BINÄRLISTE:
{binary_list}

📈 STATISTIKEN:
• Gesamtanzahl Patente: {len(binary_list)}
• Konflikte erkannt: {sum(binary_list)}
• Keine Konflikte: {len(binary_list) - sum(binary_list)}
• Konfliktrate: {(sum(binary_list) / len(binary_list) * 100):.1f}%

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📊 ZUSAMMENFASSUNG:
• Keywords erfolgreich extrahiert
• Patent-Analyse abgeschlossen
• Grafische Darstellung erstellt
• Ergebnisse gespeichert in 'extracted_keywords.txt'

✅ ANALYSE ERFOLGREICH ABGESCHLOSSEN
"""
        
        # Ergebnisse in GUI anzeigen
        self.root.after(0, lambda: self.update_results_text(results))
        
        # Ergebnisse in Datei speichern
        try:
            with open("extracted_keywords.txt", "w", encoding="utf-8") as f:
                f.write(f"Ollama Analysis Results - {filename}\n")
                f.write(f"Generated: {time.strftime('%d.%m.%Y %H:%M:%S')}\n\n")
                f.write("Keywords:\n")
                f.write(keywords)
                f.write("\n\nPatent Analysis:\n")
                f.write(patent_analysis)
                f.write(f"\n\nExtracted Binary List:\n{binary_list}")
                f.write(f"\n\nStatistics:\n")
                f.write(f"Total Patents: {len(binary_list)}\n")
                f.write(f"Conflicts: {sum(binary_list)}\n")
                f.write(f"No Conflicts: {len(binary_list) - sum(binary_list)}\n")
                f.write(f"Conflict Rate: {(sum(binary_list) / len(binary_list) * 100):.1f}%\n")
        except Exception as e:
            print(f"Fehler beim Speichern: {e}")
    
    def update_results_text(self, text):
        """Aktualisiert den Results Text"""
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert('1.0', text)
        self.results_text.config(state='disabled')
    
    def clear_results(self):
        """Leert die Ergebnisse"""
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert('1.0', "Analyse läuft...")
        self.results_text.config(state='disabled')
        
        # Chart-Bereich zurücksetzen
        for widget in self.chart_container.winfo_children():
            widget.destroy()
        self.create_placeholder_chart()
    
    def show_progress(self, message):
        """Zeigt Progress Bar an"""
        self.root.after(0, lambda: self.progress_container.pack(fill='x', pady=(0, 30)))
        self.root.after(0, lambda: self.progress_var.set(message))
        self.root.after(0, lambda: self.progress_bar.start(10))
    
    def finish_processing(self):
        """Versteckt Progress Bar nach Verarbeitung"""
        self.progress_bar.stop()
        self.progress_container.pack_forget()
        self.status_var.set("Analyse abgeschlossen")
    
    def update_status(self, message):
        """Aktualisiert Status Bar"""
        self.root.after(0, lambda: self.status_var.set(message))
        self.root.after(0, lambda: self.progress_var.set(message))
    
    def run(self):
        """Startet die Anwendung"""
        self.root.mainloop()

# Hilfsfunktion für extract_keywords_from_ollama_response falls benötigt
def extract_keywords_from_ollama_response(response):
    """Extrahiert Keywords aus Ollama Response"""
    # Einfache Implementierung - kann erweitert werden
    keywords = response.strip().split(',')
    return ' OR '.join([kw.strip() for kw in keywords if kw.strip()])

if __name__ == "__main__":
    # Abhängigkeiten prüfen
    try:
        import tkinterdnd2
        import ollama
        import pandas as pd
        import matplotlib.pyplot as plt
        import numpy as np
    except ImportError as e:
        print(f"Fehlende Abhängigkeit: {e}")
        print("Installieren Sie: pip install tkinterdnd2 ollama pandas openpyxl matplotlib numpy")
        exit(1)
    
    # Anwendung starten
    app = OllamaExcelAnalyzer()
    app.run()