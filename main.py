import os
import re
import json
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd


# Pliki konfiguracyjne obok skryptu
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RULES_FILE = os.path.join(BASE_DIR, 'rules.json')
CATEGORIES_FILE = os.path.join(BASE_DIR, 'categories.json')


def load_rules():
    try:
        if not os.path.exists(RULES_FILE):
            return []
        with open(RULES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        cleaned = []
        for item in data:
            try:
                key = str(item[0]).strip()
                # normalize keyword to lower and stripped form for reliable matching
                key_norm = normalize_text(key)
                cat = item[1] if len(item) > 1 else ""
                cleaned.append((key_norm, cat))
            except Exception:
                continue
        return cleaned
    except Exception:
        return []


def save_rules(rules):
    try:
        with open(RULES_FILE, 'w', encoding='utf-8') as f:
            json.dump([list(x) for x in rules], f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_categories():
    try:
        if not os.path.exists(CATEGORIES_FILE):
            return {}
        with open(CATEGORIES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def normalize_text(s: str) -> str:
    """Normalize text for matching: lower, remove diacritics, replace non-alnum with spaces, collapse spaces."""
    if not isinstance(s, str):
        return ""
    # lower
    s = s.lower()
    # normalize unicode (remove accents)
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    # replace any non-alphanumeric character with space
    s = re.sub(r'[^0-9a-z]+', ' ', s)
    # collapse whitespace
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# Załaduj mapę kategorii
KATEGORIE_BAZA = load_categories()


class AplikacjaKategorii:
    def __init__(self, root):
        self.root = root
        self.root.title("Sortownia Ofert v2.0")
        self.root.geometry("900x900")

        # Nagłówek
        self.label = tk.Label(root, text="Przypisz nowe kategorie do ofert", font=("Arial", 12, "bold"))
        self.label.pack(pady=10)

        # Pliki: dwa pola wyboru (wejściowy + wyjściowy) oraz przycisk startu
        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=5, fill='x', padx=10)

        tk.Label(self.file_frame, text="Plik wejściowy:").grid(row=0, column=0, sticky='w')
        self.entry_input = tk.Entry(self.file_frame, textvariable=self.input_path_var)
        self.entry_input.grid(row=0, column=1, sticky='we', padx=5)
        btn_browse_in = tk.Button(self.file_frame, text="Wybierz...", command=lambda: self.browse_input())
        btn_browse_in.grid(row=0, column=2, padx=4)

        tk.Label(self.file_frame, text="Plik wyjściowy:").grid(row=1, column=0, sticky='w')
        self.entry_output = tk.Entry(self.file_frame, textvariable=self.output_path_var)
        self.entry_output.grid(row=1, column=1, sticky='we', padx=5)
        btn_browse_out = tk.Button(self.file_frame, text="Wybierz...", command=lambda: self.browse_output())
        btn_browse_out.grid(row=1, column=2, padx=4)

        self.btn_start = tk.Button(self.file_frame, text="Rozpocznij proces", command=lambda: self.przetworz_plik(self.input_path_var.get() or None, self.output_path_var.get() or None), bg="#008CBA", fg="white", font=("Arial", 10, "bold"))
        self.btn_start.grid(row=0, column=3, rowspan=2, padx=8, sticky='ns')
        # allow the middle column to expand
        self.file_frame.columnconfigure(1, weight=1)

        # Log
        self.log_area = scrolledtext.ScrolledText(root, height=8, state='disabled', bg="#f0f0f0")
        self.log_area.pack(pady=10, padx=10, fill='both')

        # --- UI: Edycja reguł dopasowania ---
        self.rules_frame = tk.Frame(root)
        self.rules_frame.pack(padx=10, pady=5, fill='x')

        tk.Label(self.rules_frame, text="Słowo klucz:").grid(row=0, column=0, sticky='w')
        self.entry_keyword = tk.Entry(self.rules_frame)
        self.entry_keyword.grid(row=0, column=1, sticky='we', padx=5)

        tk.Label(self.rules_frame, text="Kategoria:").grid(row=1, column=0, sticky='w')
        kategori_lista = sorted(KATEGORIE_BAZA.keys()) if KATEGORIE_BAZA else []
        # Make combobox editable so user can type to filter values
        self.combo_var = tk.StringVar()
        self.combo_category = ttk.Combobox(self.rules_frame, textvariable=self.combo_var, values=kategori_lista, state='normal')
        self.combo_category.grid(row=1, column=1, sticky='we', padx=5)
        # Bind typing to filter the dropdown; don't open suggestions automatically.
        # Arrow keys will open the dropdown and allow selection.
        self.combo_category.bind('<KeyRelease>', self._on_category_keyrelease)
        self.combo_category.bind('<Return>', self._on_category_return)
        self.combo_category.bind('<Down>', self._on_category_arrow)
        self.combo_category.bind('<Up>', self._on_category_arrow)
        self.combo_category.bind('<<ComboboxSelected>>', lambda e: None)

        self.rules_frame.columnconfigure(1, weight=1)

    # Note: individual add/save/rerun buttons removed in favor of the combined 'Dodaj i dopasuj' button

        # New combined: add rule, save, and re-run (overwrite)
        self.btn_add_and_rerun = tk.Button(self.rules_frame, text="Dodaj i dopasuj", command=self.dodaj_i_dopasuj)
        self.btn_add_and_rerun.grid(row=0, column=6, rowspan=2, padx=5)

        # Lista reguł (z panel sterowania po prawej)
        self.rules_list_frame = tk.Frame(root)
        self.rules_list_frame.pack(padx=10, pady=5, fill='both', expand=True)

        self.rules_listbox = tk.Listbox(self.rules_list_frame)
        self.rules_listbox.pack(side='left', fill='both', expand=True)

        self.rules_ctrl = tk.Frame(self.rules_list_frame)
        self.rules_ctrl.pack(side='right', fill='y', padx=(6,0))

        self.rules_scroll = tk.Scrollbar(self.rules_ctrl, command=self.rules_listbox.yview)
        self.rules_scroll.pack(side='top', fill='y')
        self.rules_listbox.config(yscrollcommand=self.rules_scroll.set)

        self.btn_delete_rule = tk.Button(self.rules_ctrl, text='-', command=self.usun_regule, width=3)
        self.btn_delete_rule.pack(pady=(8,2))

        # Niedopasowane tytuły
        self.unmatched_frame = tk.Frame(root)
        self.unmatched_frame.pack(padx=10, pady=5, fill='both', expand=True)

        tk.Label(self.unmatched_frame, text="Niedopasowane tytuły (zaznacz fragment):").pack(anchor='w')
        self.unmatched_text = scrolledtext.ScrolledText(self.unmatched_frame, height=12, wrap='none')
        self.unmatched_text.pack(fill='both', expand=True)
        # Prevent typing but allow copy/select shortcuts and mouse selection.
        # Binding general Key to break blocks paste/copy; instead we intercept keys and allow Ctrl+C / Ctrl+A.
        self.unmatched_text.bind('<Key>', self._on_unmatched_key)
        # Bind explicit copy/select shortcuts to ensure they work across platforms
        self.unmatched_text.bind('<Control-c>', self._copy_unmatched)
        self.unmatched_text.bind('<Control-C>', self._copy_unmatched)
        self.unmatched_text.bind('<Control-a>', self._select_all_unmatched)
        self.unmatched_text.bind('<Control-A>', self._select_all_unmatched)
        # Right-click menu for copy
        self.unmatched_text.bind('<Button-3>', self._show_unmatched_menu)
        # When user selects text with mouse, auto-fill the keyword entry with selection
        self.unmatched_text.bind('<ButtonRelease-1>', self._on_unmatched_selection)
        # Also support keyboard selection (Shift+arrows)
        self.unmatched_text.bind('<KeyRelease>', self._on_unmatched_selection)

        #btns_frame = tk.Frame(self.unmatched_frame)
        #btns_frame.pack(fill='x', pady=4)
        #self.btn_add_from_selection = tk.Button(btns_frame, text="Dodaj regułę z zaznaczenia", command=self.dodaj_z_zaznaczenia)
        #self.btn_add_from_selection.pack(side='left')

        # pamięć
        self.last_df = None
        self.last_input_path = None
        self.last_output_path = None

        # Bottom: export final placed at bottom-right
        self.bottom_frame = tk.Frame(root)
        self.bottom_frame.pack(fill='x', side='bottom', padx=10, pady=6)
        self.btn_export_final = tk.Button(self.bottom_frame, text="Export final (id, category_id)", command=self.export_final)
        self.btn_export_final.pack(side='right')

        self.odswiez_liste_regul()

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update()

    # --- Handlers to allow copying from the read-only unmatched_text widget ---
    def _on_unmatched_key(self, event):
        """Allow only copy/select shortcuts, block other typing keys."""
        try:
            # On many platforms Control mask is bit 0x4 in event.state
            ctrl = (event.state & 0x4) != 0
        except Exception:
            ctrl = False
        if ctrl and event.keysym.lower() in ('c', 'a'):
            return None
        # Allow navigation keys so user can move the cursor/selection with arrows
        if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Next', 'Prior', 'Page_Up', 'Page_Down'):
            return None
        # Block everything else (typing, paste, delete)
        return 'break'

    def _copy_unmatched(self, event=None):
        try:
            sel = self.unmatched_text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except Exception:
            # nothing selected
            return 'break'
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(sel)
        except Exception:
            pass
        return 'break'

    def _select_all_unmatched(self, event=None):
        try:
            self.unmatched_text.tag_add(tk.SEL, '1.0', tk.END)
            self.unmatched_text.mark_set(tk.INSERT, '1.0')
            self.unmatched_text.see(tk.INSERT)
        except Exception:
            pass
        return 'break'

    def _show_unmatched_menu(self, event):
        try:
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label='Kopiuj', command=lambda: self._copy_unmatched())
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass

    def _on_unmatched_selection(self, event=None):
        """When user selects text in unmatched_text, copy it into the keyword entry box."""
        try:
            # temporarily enable to read selection
            self.unmatched_text.config(state='normal')
            sel = self.unmatched_text.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
        except Exception:
            sel = ""
        finally:
            try:
                self.unmatched_text.config(state='disabled')
            except Exception:
                pass
        if sel:
            try:
                self.entry_keyword.delete(0, tk.END)
                self.entry_keyword.insert(0, sel)
                # set focus to keyword entry so user can edit further or press tab
                self.entry_keyword.focus_set()
            except Exception:
                pass
        # Do not block the event
        return None

    # --- Handlers for the editable category combobox ---
    def _on_category_keyrelease(self, event):
        """Filter combobox values while the user types. Allow arrow/enter keys to behave normally."""
        try:
            ks = event.keysym
        except Exception:
            ks = None
        if ks in ('Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab'):
            return
        typed = self.combo_var.get()
        if not typed:
            vals = sorted(KATEGORIE_BAZA.keys()) if KATEGORIE_BAZA else []
        else:
            typed_norm = normalize_text(typed)
            vals = [k for k in KATEGORIE_BAZA.keys() if typed_norm in normalize_text(k)]
            vals = sorted(vals)
        # update dropdown values
        try:
            self.combo_category['values'] = vals
            # update values and open the dropdown so suggestions stay visible while typing
            # Note: _on_category_arrow no longer generates synthetic events, so this is safe.
            if vals:
                try:
                    # open dropdown but keep focus and caret in the entry so typing can continue
                    self.combo_category.event_generate('<Down>')
                    # ensure focus remains in the combobox entry and caret at end
                    try:
                        self.combo_category.focus_set()
                        cur = self.combo_var.get()
                        # icursor expects an index position
                        self.combo_category.icursor(len(cur))
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass

    def _on_category_return(self, event):
        """Accept current typed value or the first match when Enter is pressed."""
        cur = self.combo_var.get().strip()
        vals = list(self.combo_category['values']) if self.combo_category['values'] is not None else []
        # If exact match exists, keep it. Otherwise select first suggestion if present.
        if cur in vals:
            self.combo_var.set(cur)
        elif vals:
            self.combo_var.set(vals[0])
        return 'break'

    def _on_category_arrow(self, event):
        """Open the dropdown when user presses Up/Down to allow keyboard selection."""
        try:
            # Do nothing special here — allow default widget behavior to open/navigate the list
            pass
        except Exception:
            pass
        # Allow the event to be further processed so arrow keys navigate selections
        return None

    def znajdz_kategorie(self, tytul):
        if not isinstance(tytul, str):
            return None, None
        # normalize title the same way rules are normalized so matching is consistent
        tytul_lower = normalize_text(tytul)
        rules = load_rules()
        for slowo_klucz, nazwa_kategorii in rules:
            if slowo_klucz in tytul_lower:
                cat_id = KATEGORIE_BAZA.get(nazwa_kategorii, "")
                return nazwa_kategorii, cat_id
        return "", ""

    def przetworz_plik(self, sciezka_wejsciowa=None, sciezka_wyjsciowa=None):
        """Process input file and write output. If paths are provided, they are used; otherwise file dialogs are shown.
        """
        if not sciezka_wejsciowa:
            sciezka_wejsciowa = filedialog.askopenfilename(title="Wybierz plik do przeliczenia", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not sciezka_wejsciowa:
            return
        # show chosen file
        self.input_path_var.set(sciezka_wejsciowa)
        self.log(f"-> Pobieram: {os.path.basename(sciezka_wejsciowa)}")
        try:
            if sciezka_wejsciowa.endswith('.csv'):
                df = pd.read_csv(sciezka_wejsciowa)
            else:
                try:
                    df = pd.read_excel(sciezka_wejsciowa)
                except Exception:
                    df = pd.read_csv(sciezka_wejsciowa)
            if 'Title' not in df.columns:
                messagebox.showerror("Błąd", "Plik wejściowy nie zawiera kolumny 'Title'.")
                return
            self.log("-> Mielę dane... Czekaj.")
            nowe_kategorie = []
            nowe_ids = []
            for tytul in df['Title']:
                cat, cat_id = self.znajdz_kategorie(tytul)
                nowe_kategorie.append(cat)
                nowe_ids.append(cat_id)
            df['Category'] = nowe_kategorie
            df['Category Id'] = nowe_ids
            zmienione = len([x for x in nowe_kategorie if x])
            self.log(f"-> Zidentyfikowano {zmienione} produktów.")
            self.last_df = df
            self.odswiez_liste_regul()
            try:
                self.refresh_unmatched_list()
            except Exception:
                pass

            # if output path not provided, ask where to save
            if not sciezka_wyjsciowa:
                sciezka_wyjsciowa = filedialog.asksaveasfilename(
                    title="Gdzie zapisać gotowca?",
                    defaultextension=".xlsx",
                    filetypes=[("Plik Excel", "*.xlsx"), ("Plik CSV", "*.csv")]
                )
            if not sciezka_wyjsciowa:
                self.log("-> Anulowano zapis. A szkoda.")
                return
            # set output path in UI
            self.output_path_var.set(sciezka_wyjsciowa)
            if sciezka_wyjsciowa.endswith('.csv'):
                df.to_csv(sciezka_wyjsciowa, index=False, sep=',')
            else:
                df.to_excel(sciezka_wyjsciowa, index=False)
            self.last_input_path = sciezka_wejsciowa
            self.last_output_path = sciezka_wyjsciowa
            self.log(f"-> SUKCES! Zapisano w:\n{sciezka_wyjsciowa}")
            messagebox.showinfo("Gotowe", "Robota skończona, plik zapisany tam gdzie chciałeś.")
        except Exception as e:
            self.log(f"-> BŁĄD KRYTYCZNY: {e}")
            messagebox.showerror("Błąd", str(e))

    def browse_input(self):
        path = filedialog.askopenfilename(title="Wybierz plik do przeliczenia", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if path:
            try:
                self.input_path_var.set(path)
            except Exception:
                pass

    def browse_output(self):
        path = filedialog.asksaveasfilename(title="Gdzie zapisać gotowca?", defaultextension=".xlsx", filetypes=[("Plik Excel", "*.xlsx"), ("Plik CSV", "*.csv")])
        if path:
            try:
                self.output_path_var.set(path)
            except Exception:
                pass

    def odswiez_liste_regul(self):
        try:
            self.rules_listbox.delete(0, tk.END)
            rules = load_rules()
            for slowo, kat in rules:
                self.rules_listbox.insert(tk.END, f"{slowo} -> {kat}")
        except Exception:
            pass

    def refresh_unmatched_list(self):
        try:
            self.unmatched_text.config(state='normal')
            self.unmatched_text.delete('1.0', tk.END)
            if self.last_df is None:
                self.unmatched_text.config(state='disabled')
                return
            for idx, row in self.last_df.iterrows():
                title = str(row.get('Title', ''))
                cat = row.get('Category', '')
                if not cat:
                    self.unmatched_text.insert(tk.END, title + "\n")
            self.unmatched_text.config(state='disabled')
        except Exception:
            pass

    def dodaj_regule(self):
        key = self.entry_keyword.get().strip().lower()
        cat = self.combo_category.get().strip()
        if not key or not cat:
            messagebox.showwarning("Uwaga", "Wypełnij oba pola: słowo klucz i kategoria.")
            return
        rules = load_rules()
        rules.insert(0, (key, cat))
        saved = save_rules(rules)
        if saved:
            self.log(f"-> Dodano regułę: '{key}' -> '{cat}'")
        else:
            self.log(f"-> Błąd zapisu reguły: '{key}' -> '{cat}'")
        self.entry_keyword.delete(0, tk.END)
        try:
            self.combo_category.set("")
        except Exception:
            pass
        self.odswiez_liste_regul()

    def zapisz_reguly(self):
        rules = load_rules()
        ok = save_rules(rules)
        if ok:
            self.log("-> Reguły zapisane do rules.json")
            messagebox.showinfo("Zapisano", "Reguły zostały zapisane do pliku rules.json")
        else:
            self.log("-> Błąd zapisu reguł")
            messagebox.showerror("Błąd", "Nie udało się zapisać reguł.")

    def dodaj_i_dopasuj(self):
        """Add current rule (keyword + category), save to rules.json and immediately re-run matching (overwrite last output).
        This combines adding, saving and ponowne dopasowanie into one action.
        """
        key = self.entry_keyword.get().strip().lower()
        cat = self.combo_category.get().strip()
        if not key or not cat:
            messagebox.showwarning("Uwaga", "Wypełnij oba pola: słowo klucz i kategoria.")
            return

        # Add new rule at top
        rules = load_rules()
        rules.insert(0, (key, cat))
        saved = save_rules(rules)
        if saved:
            self.log(f"-> Dodano regułę: '{key}' -> '{cat}' (i zapisano)")
        else:
            self.log(f"-> Błąd przy zapisywaniu reguły: '{key}' -> '{cat}'")

        # clear inputs
        try:
            self.entry_keyword.delete(0, tk.END)
            self.combo_category.set("")
        except Exception:
            pass

        # If possible, re-run matching and overwrite last output
        try:
            self.ponownie_dopasuj()
        except Exception as e:
            self.log(f"-> Błąd podczas ponownego dopasowania po dodaniu reguły: {e}")

    def dodaj_z_zaznaczenia(self):
        try:
            self.unmatched_text.config(state='normal')
            fragment = self.unmatched_text.get(tk.SEL_FIRST, tk.SEL_LAST).strip().lower()
        except Exception:
            fragment = ""
        finally:
            try:
                self.unmatched_text.config(state='disabled')
            except Exception:
                pass
        if not fragment:
            messagebox.showwarning("Uwaga", "Zaznacz fragment tytułu, który chcesz dodać jako klucz.")
            return
        cat = self.combo_category.get().strip()
        if not cat:
            messagebox.showwarning("Uwaga", "Wybierz kategorię z listy.")
            return
        rules = load_rules()
        rules.insert(0, (fragment, cat))
        saved = save_rules(rules)
        if saved:
            self.log(f"-> Dodano regułę z zaznaczenia: '{fragment}' -> '{cat}'")
            self.odswiez_liste_regul()
            if self.last_df is not None:
                nowe_kategorie = []
                nowe_ids = []
                for tytul in self.last_df['Title']:
                    cat2, cat_id = self.znajdz_kategorie(tytul)
                    nowe_kategorie.append(cat2)
                    nowe_ids.append(cat_id)
                self.last_df['Category'] = nowe_kategorie
                self.last_df['Category Id'] = nowe_ids
                self.refresh_unmatched_list()
        else:
            messagebox.showerror("Błąd", "Nie udało się zapisać reguły.")

    def usun_regule(self):
        try:
            sel = self.rules_listbox.curselection()
            if not sel:
                messagebox.showwarning("Uwaga", "Wybierz regułę do usunięcia.")
                return
            idx = int(sel[0])
            rules = load_rules()
            if idx < 0 or idx >= len(rules):
                messagebox.showwarning("Uwaga", "Nieprawidłowy wybór.")
                return
            removed = rules.pop(idx)
            saved = save_rules(rules)
            if saved:
                self.log(f"-> Usunięto regułę: '{removed[0]}' -> '{removed[1]}'")
                if self.last_df is not None:
                    nowe_kategorie = []
                    nowe_ids = []
                    for tytul in self.last_df['Title']:
                        cat2, cat_id = self.znajdz_kategorie(tytul)
                        nowe_kategorie.append(cat2)
                        nowe_ids.append(cat_id)
                    self.last_df['Category'] = nowe_kategorie
                    self.last_df['Category Id'] = nowe_ids
                    self.refresh_unmatched_list()
                self.odswiez_liste_regul()
            else:
                messagebox.showerror("Błąd", "Nie udało się zapisać pliku reguł po usunięciu.")
        except Exception as e:
            messagebox.showerror("Błąd", str(e))

    def ponownie_dopasuj(self):
        if not getattr(self, 'last_input_path', None) or not getattr(self, 'last_output_path', None):
            messagebox.showwarning("Brak pliku", "Brak informacji o ostatnio przetworzonym pliku. Najpierw wykonaj przetwarzanie i zapisz wynik.")
            return
        try:
            # Ensure we use the up-to-date rules from disk
            _ = load_rules()

            sciezka_wejsciowa = self.last_input_path
            if sciezka_wejsciowa.endswith('.csv'):
                df = pd.read_csv(sciezka_wejsciowa)
            else:
                try:
                    df = pd.read_excel(sciezka_wejsciowa)
                except Exception:
                    df = pd.read_csv(sciezka_wejsciowa)
            if 'Title' not in df.columns:
                messagebox.showerror("Błąd", "Plik źródłowy nie zawiera kolumny 'Title'.")
                return
            prior_matched = 0
            try:
                if self.last_df is not None and 'Category' in self.last_df.columns:
                    prior_matched = len([x for x in self.last_df['Category'] if x])
            except Exception:
                prior_matched = 0
            nowe_kategorie = []
            nowe_ids = []
            for tytul in df['Title']:
                cat, cat_id = self.znajdz_kategorie(tytul)
                nowe_kategorie.append(cat)
                nowe_ids.append(cat_id)
            df['Category'] = nowe_kategorie
            df['Category Id'] = nowe_ids
            new_matched = len([x for x in df['Category'] if x])
            diff = new_matched - prior_matched
            sciezka_wyjsciowa = self.last_output_path
            if sciezka_wyjsciowa.endswith('.csv'):
                df.to_csv(sciezka_wyjsciowa, index=False, sep=',')
            else:
                df.to_excel(sciezka_wyjsciowa, index=False)
            self.last_df = df
            self.odswiez_liste_regul()
            self.refresh_unmatched_list()
            # Log the change: więcej/ mniej and final count
            more_text = f"+{diff}" if diff >= 0 else f"{diff}"
            status_text = "więcej" if diff > 0 else ("mniej" if diff < 0 else "bez zmian")
            self.log(f"-> PONOWNE DOPASOWANIE: zapisano nadpisany plik:\n{sciezka_wyjsciowa}\n-> Dopasowano {new_matched} ofert ({status_text}, zmiana: {more_text} względem poprzednio dopasowanych {prior_matched}).")
            #messagebox.showinfo("Gotowe", f"Dopasowano {new_matched} ofert (zmiana: {more_text} względem poprzednio dopasowanych {prior_matched}).")
        except Exception as e:
            self.log(f"-> Błąd podczas ponownego dopasowania: {e}")
            messagebox.showerror("Błąd", str(e))

    def export_final(self):
        """Export final file with only columns: id, category_id.
        Uses self.last_df if available, otherwise tries to re-read the last input file.
        """
        df = None
        try:
            if self.last_df is not None:
                df = self.last_df.copy()
            elif getattr(self, 'last_input_path', None):
                path = self.last_input_path
                if path.endswith('.csv'):
                    df = pd.read_csv(path)
                else:
                    try:
                        df = pd.read_excel(path)
                    except Exception:
                        df = pd.read_csv(path)
            else:
                messagebox.showwarning("Brak danych", "Brak przetworzonych danych. Najpierw przetwórz plik.")
                return

            if df is None:
                messagebox.showerror("Błąd", "Nie udało się załadować danych do eksportu.")
                return

            # Detect id column
            id_candidates = ['id', 'ID', 'Id', 'item_id', 'ItemId', 'itemId']
            id_col = None
            for c in id_candidates:
                if c in df.columns:
                    id_col = c
                    break
            # If still not found, try lowercase mapping
            if id_col is None:
                for c in df.columns:
                    if str(c).strip().lower() == 'id':
                        id_col = c
                        break

            # Detect category id column
            catid_candidates = ['Category Id', 'CategoryId', 'category_id', 'category id', 'Category_ID', 'categoryId']
            catid_col = None
            for c in catid_candidates:
                if c in df.columns:
                    catid_col = c
                    break
            if catid_col is None:
                for c in df.columns:
                    if str(c).strip().lower().replace(' ', '_') == 'category_id' or str(c).strip().lower().replace(' ', '') == 'categoryid':
                        catid_col = c
                        break

            # If category id missing but Category (name) present, map names to ids using KATEGORIE_BAZA
            if catid_col is None and 'Category' in df.columns:
                # create a category id column
                df['__category_id_tmp__'] = df['Category'].map(lambda x: KATEGORIE_BAZA.get(x, '') if isinstance(x, str) else '')
                catid_col = '__category_id_tmp__'

            if id_col is None:
                messagebox.showerror("Brak kolumny ID", "Nie znaleziono kolumny z identyfikatorami (id) w danych wejściowych.")
                return

            # Build output frame with normalized column names
            out_df = pd.DataFrame()
            out_df['id'] = df[id_col]
            out_df['category_id'] = df[catid_col] if catid_col in df.columns else ''

            # Remove rows without a category_id (empty or null)
            total_before = len(out_df)
            mask = out_df['category_id'].notnull() & (out_df['category_id'].astype(str).str.strip() != '')
            out_filtered = out_df[mask]
            skipped = total_before - len(out_filtered)
            if len(out_filtered) == 0:
                messagebox.showwarning("Brak danych do eksportu", "Brak wierszy z przypisanym category_id — nic nie zapisano.")
                self.log("-> Eksport przerwany: brak wierszy z category_id.")
                return

            # Ask where to save
            save_path = filedialog.asksaveasfilename(title="Gdzie zapisać finalny plik?", defaultextension='.xlsx', filetypes=[('Plik Excel', '*.xlsx'), ('Plik CSV', '*.csv')])
            if not save_path:
                self.log("-> Anulowano eksport finalny.")
                return
            if save_path.endswith('.csv'):
                out_filtered.to_csv(save_path, index=False)
            else:
                out_filtered.to_excel(save_path, index=False)

            self.log(f"-> Eksport finalny zapisany: {save_path} (wyeksportowano: {len(out_filtered)}, pominieto: {skipped})")
            messagebox.showinfo("Gotowe", f"Eksport zapisany. Wyeksportowano: {len(out_filtered)} (pominieto: {skipped}).")
        except Exception as e:
            self.log(f"-> Błąd podczas eksportu finalnego: {e}")
            messagebox.showerror("Błąd", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = AplikacjaKategorii(root)
    root.mainloop()
