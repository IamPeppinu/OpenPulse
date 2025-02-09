import os
import sqlite3
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import fitz
import shutil
from tkinter import simpledialog
import sys
from PyPDF2 import PdfReader
import pytesseract
from pdf2image import convert_from_path

base_directory = os.path.dirname(os.path.abspath(__file__))
cv_directory = os.path.join(base_directory, "CV")
if not os.path.exists(cv_directory):
    os.makedirs(cv_directory)

conn = sqlite3.connect('lavoratori.db')
c = conn.cursor()

try:
    c.execute("ALTER TABLE lavoratori ADD COLUMN cv_path TEXT")
except sqlite3.OperationalError:
    pass

try:
    c.execute("ALTER TABLE lavoratori ADD COLUMN cv_text TEXT")
except sqlite3.OperationalError:
    pass

c.execute(''' 
    CREATE TABLE IF NOT EXISTS lavoratori (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT,
        cognome TEXT,
        telefono TEXT,
        email TEXT,
        citta TEXT,
        anno_nascita INTEGER,
        valutazione TEXT,
        status TEXT,
        sesso TEXT,
        fonte_reclutamento TEXT,
        cv_path TEXT,
        cv_text TEXT
    )
''')

conn.commit()


def aggiorna_treeview():
    try:
        for row in tree.get_children():
            tree.delete(row)

        c.execute("SELECT * FROM lavoratori")
        risultati = c.fetchall()

        for row in risultati:
            tree.insert("", tk.END, values=row)
    except Exception as e:
        print("Errore nell'aggiornamento del Treeview:", e)



def inserisci_lavoratore():
    root.withdraw()
    inserimento_window = tk.Toplevel(root)
    inserimento_window.title("Inserisci Lavoratore")
    inserimento_window.configure(bg="#f0f8ff")

    def on_close():
        root.deiconify()
        inserimento_window.destroy()

    inserimento_window.protocol("WM_DELETE_WINDOW", on_close)

    tk.Label(inserimento_window, text="Nome", bg="#f0f8ff").grid(row=0, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Cognome", bg="#f0f8ff").grid(row=1, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Telefono", bg="#f0f8ff").grid(row=2, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Email", bg="#f0f8ff").grid(row=3, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Città", bg="#f0f8ff").grid(row=4, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Anno di Nascita", bg="#f0f8ff").grid(row=5, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Sesso", bg="#f0f8ff").grid(row=6, column=0, padx=5, pady=5)

    sesso_var = tk.StringVar(value="Maschio")
    sesso_menu = ttk.Combobox(inserimento_window, textvariable=sesso_var)
    sesso_menu['values'] = ("Maschio", "Femmina")
    sesso_menu.grid(row=6, column=1, padx=5, pady=5)
    sesso_menu.current(0)

    tk.Label(inserimento_window, text="Fonte Reclutamento", bg="#f0f8ff").grid(row=7, column=0, padx=5, pady=5)
    fonte_reclutamento_var = tk.StringVar(value="Indeed")
    fonte_reclutamento_menu = ttk.Combobox(inserimento_window, textvariable=fonte_reclutamento_var)
    fonte_reclutamento_menu['values'] = ("Indeed", "HelpLavoro", "Bakeca", "InfoJobs", "Linkedin", "Email", "Whatsapp", "Almalaurea", "Front Office", "Altro")
    fonte_reclutamento_menu.grid(row=7, column=1, padx=5, pady=5)
    fonte_reclutamento_menu.current(0)

    tk.Label(inserimento_window, text="Valutazione", bg="#f0f8ff").grid(row=8, column=0, padx=5, pady=5)
    tk.Label(inserimento_window, text="Status", bg="#f0f8ff").grid(row=9, column=0, padx=5, pady=5)

    frame_valutazione = tk.Frame(inserimento_window)
    frame_valutazione.grid(row=8, column=1, padx=5, pady=5)

    entry_valutazione = tk.Text(frame_valutazione, height=20, width=75, wrap=tk.WORD)
    scrollbar_valutazione = tk.Scrollbar(frame_valutazione, command=entry_valutazione.yview)
    entry_valutazione.config(yscrollcommand=scrollbar_valutazione.set)

    entry_valutazione.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar_valutazione.pack(side=tk.RIGHT, fill=tk.Y)

    entry_nome = tk.Entry(inserimento_window)
    entry_cognome = tk.Entry(inserimento_window)
    entry_telefono = tk.Entry(inserimento_window)
    entry_email = tk.Entry(inserimento_window)
    entry_citta = tk.Entry(inserimento_window)
    entry_anno_nascita = tk.Entry(inserimento_window)

    entry_nome.grid(row=0, column=1, padx=5, pady=5)
    entry_cognome.grid(row=1, column=1, padx=5, pady=5)
    entry_telefono.grid(row=2, column=1, padx=5, pady=5)
    entry_email.grid(row=3, column=1, padx=5, pady=5)
    entry_citta.grid(row=4, column=1, padx=5, pady=5)
    entry_anno_nascita.grid(row=5, column=1, padx=5, pady=5)

    status_var = tk.StringVar(value="Non Valutato")
    status_menu = ttk.Combobox(inserimento_window, textvariable=status_var)
    status_menu['values'] = ("Valutato", "Non Valutato")
    status_menu.grid(row=9, column=1, padx=5, pady=5)
    status_menu.current(1)

    cv_path = None

    def allega_cv():
        nonlocal cv_path
        cv_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if cv_path:
            messagebox.showinfo("CV Allegato", "CV allegato con successo!")

    def estrai_testo_pdf(pdf_path):
        try:
            doc = fitz.open(pdf_path)
            testo = ""
            for page in doc:
                testo += page.get_text("text")
            if testo:
                return testo
        except Exception:
            pass
        
        from PyPDF2 import PdfReader
        from pdf2image import convert_from_path
        import pytesseract
        
        try:
            reader = PdfReader(pdf_path)
            testo = ""
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    testo += page_text
            if testo:
                return testo
        except Exception:
            images = convert_from_path(pdf_path)
            extracted_text = ""
            for image in images:
                extracted_text += pytesseract.image_to_string(image)
            return extracted_text if extracted_text else "Nessun testo estratto"

    def salva_lavoratore():
        nome = entry_nome.get()
        cognome = entry_cognome.get()
        telefono = entry_telefono.get()
        email = entry_email.get()
        citta = entry_citta.get()
        anno_nascita = entry_anno_nascita.get()
        valutazione = entry_valutazione.get("1.0", tk.END).strip()
        status = status_var.get()
        sesso = sesso_var.get()
        fonte_reclutamento = fonte_reclutamento_var.get()

        try:
            c.execute("INSERT INTO lavoratori (nome, cognome, telefono, email, citta, anno_nascita, valutazione, status, sesso, fonte_reclutamento) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                      (nome, cognome, telefono, email, citta, anno_nascita, valutazione, status, sesso, fonte_reclutamento))
            conn.commit()

            if cv_path:
                lavoratore_id = c.lastrowid
                cv_filename = f"{lavoratore_id}_{os.path.basename(cv_path)}"
                nuovo_percorso = os.path.join(cv_directory, cv_filename)
                shutil.copy(cv_path, nuovo_percorso)
                c.execute("UPDATE lavoratori SET cv_path = ? WHERE id = ?", (nuovo_percorso, lavoratore_id))

                cv_text = estrai_testo_pdf(cv_path)
                c.execute("UPDATE lavoratori SET cv_text = ? WHERE id = ?", (cv_text, lavoratore_id))
                conn.commit()

            aggiorna_treeview()

            messagebox.showinfo("Successo", "Lavoratore inserito con successo!")

            root.deiconify()

        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'inserimento: {str(e)}")

        inserimento_window.destroy()

    tk.Button(inserimento_window, text="Allega CV", command=allega_cv, bg="#1E90FF", fg="black", relief=tk.RAISED).grid(row=10, columnspan=2, pady=10)
    tk.Button(inserimento_window, text="Salva Lavoratore", command=salva_lavoratore, bg="#1E90FF", fg="black", relief=tk.RAISED).grid(row=11, columnspan=2, pady=10)


def cerca_lavoratori():
    global tree  
    root.withdraw() 

    ricerca_window = tk.Toplevel(root)
    ricerca_window.title("Cerca Lavoratori")
    ricerca_window.configure(bg="#f0f8ff")
    ricerca_window.geometry("1500x900") 

    padding_x = 5  
    padding_y = 3

    def on_close():
        root.deiconify() 
        ricerca_window.destroy()

    ricerca_window.protocol("WM_DELETE_WINDOW", on_close)

    tk.Label(ricerca_window, text="Nome", bg="#f0f8ff").grid(row=0, column=0, padx=padding_x, pady=padding_y, sticky="E")
    tk.Label(ricerca_window, text="Cognome", bg="#f0f8ff").grid(row=1, column=0, padx=padding_x, pady=padding_y, sticky="E")
    tk.Label(ricerca_window, text="Telefono", bg="#f0f8ff").grid(row=2, column=0, padx=padding_x, pady=padding_y, sticky="E")
    tk.Label(ricerca_window, text="Email", bg="#f0f8ff").grid(row=3, column=0, padx=padding_x, pady=padding_y, sticky="E")
    tk.Label(ricerca_window, text="Città", bg="#f0f8ff").grid(row=4, column=0, padx=padding_x, pady=padding_y, sticky="E")
    tk.Label(ricerca_window, text="Sesso", bg="#f0f8ff").grid(row=5, column=0, padx=padding_x, pady=padding_y, sticky="E")
    sesso_var = tk.StringVar(value="")
    sesso_menu = ttk.Combobox(ricerca_window, textvariable=sesso_var, width=20)
    sesso_menu['values'] = ("Maschio", "Femmina", "")
    sesso_menu.grid(row=5, column=1, padx=padding_x, pady=padding_y, sticky="W")
    tk.Label(ricerca_window, text="Fonte Reclutamento", bg="#f0f8ff").grid(row=6, column=0, padx=padding_x, pady=padding_y, sticky="E")
    fonte_reclutamento_var = tk.StringVar(value="")
    fonte_reclutamento_menu = ttk.Combobox(ricerca_window, textvariable=fonte_reclutamento_var, width=20)
    fonte_reclutamento_menu['values'] = ("Indeed", "HelpLavoro", "Bakeca", "InfoJobs", "Linkedin", "Email", "Whatsapp", "Almalaurea", "Front Office", "Altro", "")
    fonte_reclutamento_menu.grid(row=6, column=1, padx=padding_x, pady=padding_y, sticky="W")

    tk.Label(ricerca_window, text="Anno di Nascita Da", bg="#f0f8ff").grid(row=7, column=0, padx=padding_x, pady=padding_y, sticky="E")
    tk.Label(ricerca_window, text="Anno di Nascita A", bg="#f0f8ff").grid(row=8, column=0, padx=padding_x, pady=padding_y, sticky="E")

    tk.Label(ricerca_window, text="Ricerca parole generali su valutazione e CV", bg="#f0f8ff").grid(row=9, column=0, padx=padding_x, pady=padding_y, sticky="E")
    entry_valutazione_generale = tk.Entry(ricerca_window, width=25)
    entry_valutazione_generale.grid(row=9, column=1, padx=padding_x, pady=padding_y, sticky="W")

    tk.Label(ricerca_window, text="Ricerca frase esatta su valutazione e CV", bg="#f0f8ff").grid(row=10, column=0, padx=padding_x, pady=padding_y, sticky="E")
    entry_valutazione_frase = tk.Entry(ricerca_window, width=25)
    entry_valutazione_frase.grid(row=10, column=1, padx=padding_x, pady=padding_y, sticky="W")

    tk.Label(ricerca_window, text="Escludi frase dalla valutazione", bg="#f0f8ff").grid(row=11, column=0, padx=padding_x, pady=padding_y, sticky="E")
    entry_escludi_valutazione_frase = tk.Entry(ricerca_window, width=25)
    entry_escludi_valutazione_frase.grid(row=11, column=1, padx=padding_x, pady=padding_y, sticky="W")

    tk.Label(ricerca_window, text="Status", bg="#f0f8ff").grid(row=12, column=0, padx=padding_x, pady=padding_y, sticky="E")

    entry_nome = tk.Entry(ricerca_window, width=20)
    entry_cognome = tk.Entry(ricerca_window, width=20)
    entry_telefono = tk.Entry(ricerca_window, width=20)
    entry_email = tk.Entry(ricerca_window, width=20)
    entry_citta = tk.Entry(ricerca_window, width=20)
    entry_anno_nascita_da = tk.Entry(ricerca_window, width=20)
    entry_anno_nascita_a = tk.Entry(ricerca_window, width=20)

    entry_nome.grid(row=0, column=1, padx=padding_x, pady=padding_y, sticky="W")
    entry_cognome.grid(row=1, column=1, padx=padding_x, pady=padding_y, sticky="W")
    entry_telefono.grid(row=2, column=1, padx=padding_x, pady=padding_y, sticky="W")
    entry_email.grid(row=3, column=1, padx=padding_x, pady=padding_y, sticky="W")
    entry_citta.grid(row=4, column=1, padx=padding_x, pady=padding_y, sticky="W")
    entry_anno_nascita_da.grid(row=7, column=1, padx=padding_x, pady=padding_y, sticky="W")
    entry_anno_nascita_a.grid(row=8, column=1, padx=padding_x, pady=padding_y, sticky="W")

    status_var = tk.StringVar(value="")
    status_menu = ttk.Combobox(ricerca_window, textvariable=status_var, width=20)
    status_menu['values'] = ("Valutato", "Non Valutato", "")
    status_menu.grid(row=12, column=1, padx=padding_x, pady=padding_y, sticky="W")

    label_numero_risultati = tk.Label(ricerca_window, text="Numero risultati: 0", bg="#f0f8ff")
    label_numero_risultati.grid(row=13, columnspan=2, pady=padding_y)

    def esegui_ricerca():

        query = "SELECT * FROM lavoratori WHERE 1=1"
        params = []

        if entry_nome.get():
            query += " AND nome LIKE ?"
            params.append(f"%{entry_nome.get()}%")
        if entry_cognome.get():
            query += " AND cognome LIKE ?"
            params.append(f"%{entry_cognome.get()}%")
        if entry_telefono.get():
            telefono = entry_telefono.get().replace(" ", "")  
            query += " AND REPLACE(telefono, ' ', '') LIKE ?"
            params.append(f"%{telefono}%")
        if entry_email.get():
            query += " AND email LIKE ?"
            params.append(f"%{entry_email.get()}%")
        if entry_citta.get():
            query += " AND citta LIKE ?"
            params.append(f"%{entry_citta.get()}%")
        if entry_anno_nascita_da.get() and entry_anno_nascita_a.get():
            query += " AND anno_nascita BETWEEN ? AND ?"
            params.append(entry_anno_nascita_da.get())
            params.append(entry_anno_nascita_a.get())

        if fonte_reclutamento_var.get():
            query += " AND fonte_reclutamento = ?"
            params.append(fonte_reclutamento_var.get())

        if entry_valutazione_generale.get():
            valutazione_terms = entry_valutazione_generale.get().split()
            or_conditions = " OR ".join(["valutazione LIKE ? OR cv_text LIKE ?"] * len(valutazione_terms))
            query += f" AND ({or_conditions})"
            for term in valutazione_terms:
                params.append(f"%{term}%")
                params.append(f"%{term}%")

        if entry_valutazione_frase.get():
            frasi = [frase.strip() for frase in entry_valutazione_frase.get().split(",")]
            or_conditions = " OR ".join(["valutazione LIKE ? OR cv_text LIKE ?"] * len(frasi))
            query += f" AND ({or_conditions})"
            for frase in frasi:
                params.append(f"%{frase}%")
                params.append(f"%{frase}%")

        if entry_escludi_valutazione_frase.get():
            frasi_da_escludere = [frase.strip() for frase in entry_escludi_valutazione_frase.get().split(",")]
            for frase in frasi_da_escludere:
                query += " AND (valutazione NOT LIKE ? AND cv_text NOT LIKE ?)"
                params.append(f"%{frase}%")
                params.append(f"%{frase}%")

        if status_var.get():
            query += " AND status = ?"
            params.append(status_var.get())

        if sesso_var.get():
            query += " AND sesso = ?"
            params.append(sesso_var.get())

        c.execute(query, params)
        risultati = c.fetchall()

        for row in tree.get_children():
            tree.delete(row)
        for row in risultati:
            tree.insert("", tk.END, values=row)

        label_numero_risultati.config(text=f"Numero risultati: {len(risultati)}")

    tk.Button(ricerca_window, text="Cerca", command=esegui_ricerca, bg="#1E90FF", fg="black", relief=tk.RAISED).grid(row=14, columnspan=2, pady=10)

    columns = ("ID", "Nome", "Cognome", "Telefono", "Email", "Città", "Anno di Nascita", "Valutazione", "Status", "Sesso", "Fonte Reclutamento")
    tree = ttk.Treeview(ricerca_window, columns=columns, show="headings", height=25)

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center") 

    tree.grid(row=15, columnspan=2, pady=10)

    scrollbar = tk.Scrollbar(ricerca_window, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.grid(row=15, column=2, sticky='ns')
    
    tree.bind("<Double-1>", modifica_lavoratore)

    tk.Button(ricerca_window, text="Cancella Lavoratore", command=cancella_lavoratore, bg="#1E90FF", fg="black", relief=tk.RAISED).grid(row=14, column=0, padx=padding_x, pady=10, sticky="W")


def modifica_lavoratore(event):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Seleziona Lavoratore", "Devi selezionare un lavoratore da modificare.")
        return

    item_values = tree.item(selected_item[0], "values")
    modifica_window = tk.Toplevel(root)
    modifica_window.title("Modifica Lavoratore")
    modifica_window.configure(bg="#f0f8ff")

    tk.Label(modifica_window, text="Nome", bg="#f0f8ff").grid(row=0, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Cognome", bg="#f0f8ff").grid(row=1, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Telefono", bg="#f0f8ff").grid(row=2, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Email", bg="#f0f8ff").grid(row=3, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Città", bg="#f0f8ff").grid(row=4, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Anno di Nascita", bg="#f0f8ff").grid(row=5, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Sesso", bg="#f0f8ff").grid(row=6, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Fonte Reclutamento", bg="#f0f8ff").grid(row=7, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Valutazione", bg="#f0f8ff").grid(row=8, column=0, padx=5, pady=5)
    tk.Label(modifica_window, text="Status", bg="#f0f8ff").grid(row=9, column=0, padx=5, pady=5)

    sesso_var = tk.StringVar()
    sesso_menu = ttk.Combobox(modifica_window, textvariable=sesso_var)
    sesso_menu['values'] = ("Maschio", "Femmina", "")
    sesso_menu.grid(row=6, column=1, padx=5, pady=5)

    fonte_reclutamento_var = tk.StringVar()
    fonte_reclutamento_menu = ttk.Combobox(modifica_window, textvariable=fonte_reclutamento_var)
    fonte_reclutamento_menu['values'] = ("Indeed", "HelpLavoro", "Bakeca", "InfoJobs", "Linkedin", "Email", "Whatsapp", "Almalaurea", "Front Office", "Altro", "")
    fonte_reclutamento_menu.grid(row=7, column=1, padx=5, pady=5)

    frame_valutazione = tk.Frame(modifica_window)
    frame_valutazione.grid(row=8, column=1, padx=5, pady=5)

    entry_valutazione = tk.Text(frame_valutazione, height=20, width=75, wrap=tk.WORD)
    scrollbar_valutazione = tk.Scrollbar(frame_valutazione, command=entry_valutazione.yview)
    entry_valutazione.config(yscrollcommand=scrollbar_valutazione.set)
    entry_valutazione.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar_valutazione.pack(side=tk.RIGHT, fill=tk.Y)

    entry_nome = tk.Entry(modifica_window)
    entry_cognome = tk.Entry(modifica_window)
    entry_telefono = tk.Entry(modifica_window)
    entry_email = tk.Entry(modifica_window)
    entry_citta = tk.Entry(modifica_window)
    entry_anno_nascita = tk.Entry(modifica_window)

    status_var = tk.StringVar()
    status_menu = ttk.Combobox(modifica_window, textvariable=status_var)
    status_menu['values'] = ("Valutato", "Non Valutato")

    entry_nome.grid(row=0, column=1, padx=5, pady=5)
    entry_cognome.grid(row=1, column=1, padx=5, pady=5)
    entry_telefono.grid(row=2, column=1, padx=5, pady=5)
    entry_email.grid(row=3, column=1, padx=5, pady=5)
    entry_citta.grid(row=4, column=1, padx=5, pady=5)
    entry_anno_nascita.grid(row=5, column=1, padx=5, pady=5)
    status_menu.grid(row=9, column=1, padx=5, pady=5)

    entry_nome.insert(0, item_values[1])
    entry_cognome.insert(0, item_values[2])
    entry_telefono.insert(0, item_values[3])
    entry_email.insert(0, item_values[4])
    entry_citta.insert(0, item_values[5])
    entry_anno_nascita.insert(0, item_values[6])
    entry_valutazione.insert("1.0", item_values[7])
    sesso_var.set(item_values[9])
    fonte_reclutamento_var.set(item_values[10])
    status_var.set(item_values[8])

    def gestisci_cv():
        gestisci_window = tk.Toplevel(modifica_window)
        gestisci_window.title("Gestisci CV")
        gestisci_window.configure(bg="#f0f8ff")

        base_directory = os.path.dirname(os.path.abspath(__file__))
        cv_directory = os.path.join(base_directory, "CV")
        lavoratore_id = item_values[0]  
        cv_files = [file for file in os.listdir(cv_directory) if file.startswith(f"{lavoratore_id}_")]

        cv_listbox = tk.Listbox(gestisci_window, selectmode=tk.SINGLE, height=10, width=50)
        for file in cv_files:
            cv_listbox.insert(tk.END, file)
        cv_listbox.grid(row=0, column=0, padx=10, pady=10)

        def visualizza_cv():
            selected_cv = cv_listbox.curselection()
            if selected_cv:
                cv_file = cv_listbox.get(selected_cv[0])
                cv_path = os.path.join(cv_directory, cv_file)
                try:
                    if os.name == 'nt':
                        os.startfile(cv_path)
                    elif os.name == 'posix':
                        subprocess.call(['open', cv_path]) if sys.platform == 'darwin' else subprocess.call(['xdg-open', cv_path])
                except Exception as e:
                    messagebox.showerror("Errore", f"Impossibile aprire il file PDF: {str(e)}")

        def cancella_cv():
            selected_cv = cv_listbox.curselection()
            if selected_cv:
                cv_file = cv_listbox.get(selected_cv[0])
                cv_path = os.path.join(cv_directory, cv_file)
                os.remove(cv_path)
                cv_listbox.delete(selected_cv[0])

        def inserisci_nuovo_cv():
            file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if file_path:
                cv_file_name = f"{lavoratore_id}_{os.path.basename(file_path)}"
                new_cv_path = os.path.join(cv_directory, cv_file_name)
                shutil.copy(file_path, new_cv_path)
                cv_listbox.insert(tk.END, cv_file_name)
                open_cv_and_extract_text(new_cv_path)

        def visualizza_testo_cv():
            selected_cv = cv_listbox.curselection()
            if selected_cv:
                cv_file = cv_listbox.get(selected_cv[0])
                cv_path = os.path.join(cv_directory, cv_file)
                extracted_text = open_cv_and_extract_text(cv_path)
                messagebox.showinfo("Testo Estratto", extracted_text)

        tk.Button(gestisci_window, text="Visualizza CV", command=visualizza_cv, bg="#1E90FF", fg="black").grid(row=1, column=1, padx=5, pady=5)
        tk.Button(gestisci_window, text="Cancella CV", command=cancella_cv, bg="#1E90FF", fg="black").grid(row=2, column=1, padx=5, pady=5)
        tk.Button(gestisci_window, text="Inserisci Nuovo CV", command=inserisci_nuovo_cv, bg="#1E90FF", fg="black").grid(row=3, column=1, padx=5, pady=5)
        tk.Button(gestisci_window, text="Visualizza Testo CV", command=visualizza_testo_cv, bg="#1E90FF", fg="black").grid(row=4, column=1, padx=5, pady=5)

    tk.Button(modifica_window, text="Gestisci CV", command=gestisci_cv, bg="#1E90FF", fg="black").grid(row=10, columnspan=2, pady=10)

    def salva_modifiche():
        nome = entry_nome.get()
        cognome = entry_cognome.get()
        telefono = entry_telefono.get()
        email = entry_email.get()
        citta = entry_citta.get()
        anno_nascita = entry_anno_nascita.get()
        valutazione = entry_valutazione.get("1.0", tk.END).strip()
        status = status_var.get()
        sesso = sesso_var.get()
        fonte_reclutamento = fonte_reclutamento_var.get()

        c.execute("UPDATE lavoratori SET nome=?, cognome=?, telefono=?, email=?, citta=?, anno_nascita=?, valutazione=?, status=?, sesso=?, fonte_reclutamento=? WHERE id=?",
                  (nome, cognome, telefono, email, citta, anno_nascita, valutazione, status, sesso, fonte_reclutamento, item_values[0]))

        cv_text = get_cv_text_from_files(item_values[0])
        c.execute("UPDATE lavoratori SET cv_text=? WHERE id=?", (cv_text, item_values[0]))
        conn.commit()

        messagebox.showinfo("Successo", "Lavoratore aggiornato con successo!")
        aggiorna_treeview()
        modifica_window.destroy()

    tk.Button(modifica_window, text="Salva Modifiche", command=salva_modifiche, bg="#1E90FF", fg="black", relief=tk.RAISED).grid(row=11, columnspan=2, pady=10)

def get_cv_text_from_files(lavoratore_id):
    base_directory = os.path.dirname(os.path.abspath(__file__))
    cv_directory = os.path.join(base_directory, "CV")
    cv_files = [file for file in os.listdir(cv_directory) if file.startswith(f"{lavoratore_id}_")]

    all_text = ""
    for file in cv_files:
        cv_path = os.path.join(cv_directory, file)
        extracted_text = open_cv_and_extract_text(cv_path)
        all_text += extracted_text + "\n"

    return all_text.strip()

def open_cv_and_extract_text(cv_path):
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(cv_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        if text.strip() == "":
            text = extract_text_with_pymupdf(cv_path)
        return text
    except Exception:
        text = extract_text_with_pymupdf(cv_path)
        if text.strip() == "":
            text = extract_text_with_tesseract(cv_path)
        return text

def extract_text_with_pymupdf(cv_path):
    try:
        import fitz
        doc = fitz.open(cv_path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    except Exception:
        return ""

def extract_text_with_tesseract(cv_path):
    from pdf2image import convert_from_path
    import pytesseract

    images = convert_from_path(cv_path)
    text = ""
    for image in images:
        text += pytesseract.image_to_string(image)
    return text


def cancella_lavoratore():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Seleziona Lavoratore", "Devi selezionare un lavoratore da cancellare.")
        return
    
    item_values = tree.item(selected_item[0], "values")
    c.execute("DELETE FROM lavoratori WHERE id=?", (item_values[0],))
    conn.commit()
    
    messagebox.showinfo("Successo", "Lavoratore cancellato con successo!")
    aggiorna_treeview() 


root = tk.Tk()
root.title("OpenPulse")
root.geometry("800x600")
root.configure(bg="#f0f8ff")

tk.Button(root, text="Inserisci Lavoratore", command=inserisci_lavoratore, bg="#1E90FF", fg="black", relief=tk.RAISED).pack(pady=10)
tk.Button(root, text="Cerca Lavoratori", command=cerca_lavoratori, bg="#1E90FF", fg="black", relief=tk.RAISED).pack(pady=10)



tk.Label(root, text="Sviluppato da Giuseppe Sergio.", bg="#f0f8ff", font=("Arial", 10)).pack(side=tk.BOTTOM, pady=3)
tk.Label(root, text="OpenPulse® V2.0", bg="#f0f8ff", font=("Arial", 8)).pack(side=tk.BOTTOM, pady=3)

root.mainloop()

conn.close()
