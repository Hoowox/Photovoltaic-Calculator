import tkinter as tk
import datetime
from tkinter import ttk
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import glob
from tkinter import filedialog
import math
from PIL import Image, ImageTk
import os
import json

class CustomDialog(tk.Toplevel):
    def __init__(self, parent, title=None, message=None, initialvalue=None, texts=None):
        tk.Toplevel.__init__(self, parent)
        self.title(title)
        self.geometry('200x120')  # Zmniejszenie rozmiaru okna
        self.resizable(False, False)  # Uniemożliwienie zmiany rozmiaru okna
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.result = None
        self.texts = texts

        # Zastosowanie stylu 'TLabel' dla etykiety
        label = ttk.Label(self, text=message)
        label.pack(padx=10, pady=10)

        # Zastosowanie stylu 'TEntry' dla pola wprowadzania
        self.entry = ttk.Entry(self)
        self.entry.pack(fill='both', expand=True, padx=20, pady=5)  # Zmniejszenie pola do wpisywania
        self.entry.insert(0, initialvalue)

        buttons = ttk.Frame(self)
        buttons.pack(pady=5)  # Dodanie odstępu od dolnej krawędzi okna

        # Zastosowanie stylu 'TButton' dla przycisków
        ok_button = ttk.Button(buttons, text=self.texts['save'], command=self.ok, width=10)
        ok_button.pack(side='left', padx=5, pady=5)  # Zwiększenie wysokości przycisku

        cancel_button = ttk.Button(buttons, text=self.texts['cancel_settings'], command=self.cancel, width=10)
        cancel_button.pack(side='right', padx=5, pady=5)  # Zwiększenie wysokości przycisku

    def ok(self, event=None):
        self.result = self.entry.get()
        self.destroy()

    def cancel(self, event=None):
        self.destroy()

def askstring(title, message, initialvalue=None, texts=None):
    root = tk.Tk()
    root.withdraw()
    d = CustomDialog(root, title, message, initialvalue, texts)
    root.wait_window(d)
    return d.result

class MojaAplikacja:
    def __init__(self, root):
        self.root = root
        self.naswietlenie_miesieczne = {}
        self.temperatura_miesieczna = {}
        self.language = 'polski'
        self.texts = self.load_language(self.language)

        # Dodanie nazwy programu do tytułu okna
        self.root.title(self.texts['title'])

        # Dodanie obsługi klawisza F1
        self.root.bind('<F1>', self.pokaz_pomoc)

        # Tworzenie paska menu
        self.menu_bar = tk.Menu(root)
        self.root.config(menu=self.menu_bar)

        # Dodawanie opcji ustawień naświetlenia i temperatury
        self.settings_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.texts['settings'], menu=self.settings_menu)
        self.settings_menu.add_command(label=self.texts['monthly_insolation'], command=self.ustaw_naswietlenie)
        self.settings_menu.add_command(label=self.texts['monthly_temperatures'], command=self.ustaw_temperatury)
        language_menu = tk.Menu(self.settings_menu, tearoff=0)
        self.settings_menu.add_cascade(label=self.texts['change_language'], menu=language_menu)
        language_menu.add_command(label="Polski", command=lambda: self.change_language('polski'))
        language_menu.add_command(label="English", command=lambda: self.change_language('english'))

        # Dodawanie opcji zapisywania wykresu, ustawień i danych
        self.save_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.texts['save'], menu=self.save_menu)
        self.save_menu.add_command(label=self.texts['save_chart'], command=self.zapisz_wykres)
        self.save_menu.add_command(label=self.texts['save_settings'], command=self.zapisz_ustawienia)
        self.save_menu.add_command(label=self.texts['save_data'], command=self.zapisz_dane)

        # Dodawanie opcji importowania wykresu, ustawień i danych
        self.import_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.texts['import'], menu=self.import_menu)
        self.import_menu.add_command(label=self.texts['import_chart'], command=self.importuj_wykres)
        self.import_menu.add_command(label=self.texts['import_settings'], command=self.wybierz_i_importuj_ustawienia)
        self.import_menu.add_command(label=self.texts['import_data'], command=self.wybierz_i_importuj_dane)

        # Dodawanie opcji Pomoc
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.texts['help'], menu=self.help_menu)
        self.help_menu.add_command(label=self.texts['how_to_use'], command=self.pokaz_pomoc)

        # Dodawanie opcji Inne
        other_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.texts['other'], menu=other_menu)

        # Dodawanie opcji wyjścia z programu
        self.exit_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.texts['exit'], menu=self.exit_menu)
        self.exit_menu.add_command(label=self.texts['exit'], command=self.wyjdz)

        # Dodanie wykresu
        self.fig = plt.Figure(figsize=(6, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=root)
        self.canvas.get_tk_widget().grid(row=0, column=1, sticky='nsew')

        # Rozciągnięcie kolumny z wykresem
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(0, weight=1)

        # Tworzenie ramki na ustawienia wejściowe
        input_frame = ttk.Frame(root)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky='ns')

        # Etykiety i pola do wprowadzania danych
        labels = ["Powierzchnia (m²)", "Kąt nachylenia (stopnie)", "Azymut (stopnie)", "Moc paneli (kWp)", "Sprawność paneli (%)",
                  "Temperatura otoczenia (°C)", "Liczba godzin słonecznych", "Albedo (%)", "Zacienienie (%)",
                  "Zanieczyszczenia (%)"]

        self.labels = {}
        self.entries = {}
        for i, label in enumerate(labels):
            lbl = ttk.Label(input_frame, text=label)
            lbl.grid(row=i, column=0, sticky='w')
            self.labels[label] = lbl
            entry = ttk.Entry(input_frame)
            entry.grid(row=i, column=1)
            self.entries[label] = entry

        # Przycisk do obliczeń
        self.calculate_button = ttk.Button(input_frame, text=self.texts['calculate'], command=self.oblicz)
        self.calculate_button.grid(row=len(labels), column=0, columnspan=2, pady=10)

        # Etykieta na wynik
        self.result_label = ttk.Label(root, text=self.texts['result'])
        self.result_label.grid(row=1, column=0, padx=10, pady=10)

    def load_language(self, lang):
        with open(f'{lang}.json', 'r', encoding='utf-8') as f:
            return json.load(f)

    def update_ui_texts(self):
        # Aktualizacja tekstów w menu
        self.settings_menu.entryconfig(0, label=self.texts['monthly_insolation'])
        self.settings_menu.entryconfig(1, label=self.texts['monthly_temperatures'])
        self.settings_menu.entryconfig(2, label=self.texts['change_language'])
        self.save_menu.entryconfig(0, label=self.texts['save_chart'])
        self.save_menu.entryconfig(1, label=self.texts['save_settings'])
        self.save_menu.entryconfig(2, label=self.texts['save_data'])
        self.import_menu.entryconfig(0, label=self.texts['import_chart'])
        self.import_menu.entryconfig(1, label=self.texts['import_settings'])
        self.import_menu.entryconfig(2, label=self.texts['import_data'])
        self.help_menu.entryconfig(0, label=self.texts['how_to_use'])
        self.exit_menu.entryconfig(0, label=self.texts['exit'])
        self.menu_bar.entryconfig(1, label=self.texts['settings'])
        self.menu_bar.entryconfig(2, label=self.texts['save'])
        self.menu_bar.entryconfig(3, label=self.texts['import'])
        self.menu_bar.entryconfig(4, label=self.texts['help'])
        self.menu_bar.entryconfig(5, label=self.texts['other'])
        self.menu_bar.entryconfig(6, label=self.texts['exit'])

        # Aktualizacja tekstu przycisku i etykiety wyniku
        self.calculate_button['text'] = self.texts['calculate']
        self.result_label['text'] = self.texts['result']

        # Aktualizacja tytułu okna
        self.root.title(self.texts.get('title', ''))

        # Aktualizacja etykiet pól wprowadzania danych
        for label, lbl in self.labels.items():
            lbl['text'] = self.texts.get(label, label)

        # Aktualizacja tytułów okien temperatury i naświetlenia, jeśli są otwarte
        if hasattr(self, 'temperatura_window') and self.temperatura_window.winfo_exists():
            self.temperatura_window.title(self.texts['monthly_temperatures'])
        if hasattr(self, 'naswietlenie_window') and self.naswietlenie_window.winfo_exists():
            self.naswietlenie_window.title(self.texts['monthly_insolation'])

        # Aktualizacja komunikatu o błędzie przy wprowadzaniu złych danych
        if hasattr(self, 'result_label') and self.result_label['text'] == "Proszę wprowadzić temperaturę otoczenia.":
            self.result_label['text'] = self.texts['enter_environment_temperature']

        # Aktualizacja etykiety osi Y na wykresie
        if hasattr(self, 'ax'):
            self.ax.set_ylabel(self.texts['generated_energy_kWh'])

        # Aktualizacja komunikatu o usuwaniu naświetleń oraz temperatur
        if hasattr(self, 'naswietlenie_window') and self.naswietlenie_window.winfo_exists():
            self.naswietlenie_window.title(self.texts['delete_all_data'])

        if hasattr(self, 'temperatura_window') and self.temperatura_window.winfo_exists():
            self.temperatura_window.title(self.texts['delete_all_data'])

        # Aktualizacja tekstów związanych z zapisywaniem i importem
        self.file_name_prompt = self.texts['file_name_prompt']
        self.enter_file_name = self.texts['enter_file_name']
        self.info = self.texts['info']
        self.chart_saved = self.texts['chart_saved']
        self.settings_saved = self.texts['settings_saved']
        self.data_saved = self.texts['data_saved']
        self.chart_imported = self.texts['chart_imported']
        self.wrong_number_of_columns = self.texts['wrong_number_of_columns']
        self.missing_expected_column = self.texts['missing_expected_column']
        self.select_and_import_settings = self.texts['select_and_import_settings']
        self.select_and_import_data = self.texts['select_and_import_data']
        self.in_this_folder = self.texts['in_this_folder']

        # Aktualizacja tytułu okna pomocy
        self.help_title = self.texts['help']

        # Aktualizacja tekstu wyświetlanego, gdy plik pomocy nie zostanie znaleziony
        self.help_not_found = self.texts['help_not_found']

        # Aktualizacja tekstu wyświetlanego w oknie dialogowym przy wyjściu
        self.exit_text = self.texts['exit']
        self.exit_confirmation = self.texts['exit_confirmation']

        # Aktualizacja tekstu związanych z oknem dialogowym
        self.save = self.texts['save']
        self.cancel_settings = self.texts['cancel_settings']

    def change_language(self, lang):
        self.language = lang
        self.texts = self.load_language(self.language)
        self.update_ui_texts()

    def ustaw_naswietlenie(self):
        # Okno do ustawiania naświetlenia miesięcznego
        self.naswietlenie_window = tk.Toplevel(self.root)
        self.naswietlenie_window.title(self.texts['monthly_insolation'])

        # Ustawienie konkretnej wielkości okna i zablokowanie możliwości zmiany rozmiaru
        self.naswietlenie_window.geometry("280x380")
        self.naswietlenie_window.resizable(False, False)

        # Utworzenie ramki do wyśrodkowania elementów
        frame = tk.Frame(self.naswietlenie_window)
        frame.place(relx=0.5, rely=0.5, anchor='center')

        miesiace = self.texts['months']
        self.naswietlenie_entries = {}
        for i, miesiac in enumerate(miesiace):
            ttk.Label(frame, text=miesiac).grid(row=i, column=0, sticky='w')
            entry = ttk.Entry(frame)
            entry.grid(row=i, column=1)
            self.naswietlenie_entries[i + 1] = entry  # zmieniamy klucz na numer miesiąca

            # Wypełnienie pól wprowadzania danymi, jeśli są dostępne
            if i + 1 in self.naswietlenie_miesieczne:
                entry.insert(0, self.naswietlenie_miesieczne[i + 1])

        button_frame = tk.Frame(frame)
        button_frame.grid(row=len(miesiace), column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text=self.texts['save'], command=self.zapisz_naswietlenie).pack(side='left', padx=10)
        ttk.Button(button_frame, text=self.texts['delete'], command=self.kasuj_naswietlenie).pack(side='right', padx=10)

    def kasuj_naswietlenie(self):
        if messagebox.askyesno(self.texts['confirmation'], self.texts['delete_all_data'], parent=self.naswietlenie_window):
            self.naswietlenie_miesieczne.clear()
            for entry in self.naswietlenie_entries.values():
                entry.delete(0, tk.END)

    def ustaw_temperatury(self):
        # Okno do ustawiania temperatury otoczenia dla każdego miesiąca
        self.temperatura_window = tk.Toplevel(self.root)
        self.temperatura_window.title(self.texts['monthly_temperatures'])

        # Ustawienie konkretnej wielkości okna i zablokowanie możliwości zmiany rozmiaru
        self.temperatura_window.geometry("280x380")
        self.temperatura_window.resizable(False, False)

        # Utworzenie ramki do wyśrodkowania elementów
        frame = tk.Frame(self.temperatura_window)
        frame.place(relx=0.5, rely=0.5, anchor='center')

        miesiace = self.texts['months']
        self.temperatura_entries = {}
        for i, miesiac in enumerate(miesiace):
            ttk.Label(frame, text=miesiac).grid(row=i, column=0, sticky='w')
            entry = ttk.Entry(frame)
            entry.grid(row=i, column=1)
            self.temperatura_entries[i + 1] = entry  # zmieniamy klucz na numer miesiąca

            # Wypełnienie pól wprowadzania danymi, jeśli są dostępne
            if i + 1 in self.temperatura_miesieczna:
                entry.insert(0, self.temperatura_miesieczna[i + 1])

        button_frame = tk.Frame(frame)
        button_frame.grid(row=len(miesiace), column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text=self.texts['save'], command=self.zapisz_temperatury).pack(side='left', padx=10)
        ttk.Button(button_frame, text=self.texts['delete'], command=self.kasuj_temperatury).pack(side='right', padx=10)

    def kasuj_temperatury(self):
        if messagebox.askyesno(self.texts['confirmation'], self.texts['delete_all_data'], parent=self.temperatura_window):
            self.temperatura_miesieczna.clear()
            for entry in self.temperatura_entries.values():
                entry.delete(0, tk.END)

    def zapisz_temperatury(self):
        try:
            for miesiac, entry in self.temperatura_entries.items():
                self.temperatura_miesieczna[miesiac] = float(entry.get())
            messagebox.showinfo(self.texts['success'], self.texts['monthly_temperatures_saved'])
            self.temperatura_window.destroy()
        except ValueError:
            messagebox.showerror(self.texts['error'], self.texts['enter_valid_numeric_values'])

    def zapisz_naswietlenie(self):
        try:
            for miesiac, entry in self.naswietlenie_entries.items():
                self.naswietlenie_miesieczne[miesiac] = float(entry.get())
            messagebox.showinfo(self.texts['success'], self.texts['monthly_illumination_saved'])
            self.naswietlenie_window.destroy()
        except ValueError:
            messagebox.showerror(self.texts['error'], self.texts['enter_valid_numeric_values'])

    def oblicz_wytworzona_energie(self, powierzchnia, kat_nachylenia, azymut, moc_paneli, sprawność_paneli,
                                  temperatura_otoczenia, intensywność_światła, liczba_godzin_słonecznych, albedo,
                                  zacienienie, zanieczyszczenia):
        stała_korekcyjna = 0.75
        PR = sprawność_paneli * (1 - 0.005 * (temperatura_otoczenia - 25)) * (1 - zacienienie / 100) * (
                1 - zanieczyszczenia / 100)
        G = intensywność_światła * (liczba_godzin_słonecznych / 365) * (1 + albedo / 100)

        # Uwzględnianie kąta nachylenia i azymutu
        kat_nachylenia_rad = math.radians(kat_nachylenia)
        azymut_rad = math.radians(azymut)
        G *= math.sin(kat_nachylenia_rad) * math.cos(azymut_rad)

        # Uwzględnianie mocy paneli
        wytworzona_energia = powierzchnia * G * PR * stała_korekcyjna * moc_paneli
        return wytworzona_energia

    def oblicz(self):
        try:
            # Lista miesięcy
            miesiace = self.texts['months']

            # Pobieranie aktualnego miesiąca
            aktualny_miesiac = datetime.datetime.now().month

            # Pobieranie danych z pól
            powierzchnia = float(self.entries["Powierzchnia (m²)"].get())
            if not 1 <= powierzchnia <= 1000:
                raise ValueError('invalid_surface_area')

            kat_nachylenia = float(self.entries["Kąt nachylenia (stopnie)"].get())
            if not 0 <= kat_nachylenia <= 90:
                raise ValueError('invalid_inclination_angle')

            azymut = float(self.entries["Azymut (stopnie)"].get())
            if not -180 <= azymut <= 180:
                raise ValueError('invalid_azimuth')

            moc_paneli = float(self.entries["Moc paneli (kWp)"].get())
            if not 0.1 <= moc_paneli <= 100:
                raise ValueError('invalid_panel_power')

            sprawność_paneli = float(self.entries["Sprawność paneli (%)"].get()) / 100
            if not 0.01 <= sprawność_paneli <= 1:
                raise ValueError('invalid_panel_efficiency')

            liczba_godzin_słonecznych = float(self.entries["Liczba godzin słonecznych"].get())
            if not 0 <= liczba_godzin_słonecznych <= 8784:
                raise ValueError('invalid_sunlight_hours')

            albedo = float(self.entries["Albedo (%)"].get())
            if not 0.1 <= albedo <= 1:
                raise ValueError('invalid_albedo')

            zacienienie = float(self.entries["Zacienienie (%)"].get())
            if not 0 <= zacienienie <= 100:
                raise ValueError('invalid_shading')

            zanieczyszczenia = float(self.entries["Zanieczyszczenia (%)"].get())
            if not 0 <= zanieczyszczenia <= 100:
                raise ValueError('invalid_pollution')

            # Wartości domyślne naświetlenia dla każdego miesiąca
            domyslne_naswietlenie = {
                1: 50, 2: 75, 3: 100, 4: 125,
                5: 150, 6: 175, 7: 200, 8: 175,
                9: 150, 10: 125, 11: 100, 12: 75
            }

            # Pobieranie naświetlenia z domyślnymi wartościami, jeśli nie podano
            intensywność_światła = self.naswietlenie_miesieczne.get(aktualny_miesiac,
                                                                    domyslne_naswietlenie[aktualny_miesiac])

            # Pobieranie temperatury otoczenia z pola "Temperatura otoczenia", jeśli jest dostępna
            temperatura_otoczenia_entry = self.entries["Temperatura otoczenia (°C)"].get()
            if temperatura_otoczenia_entry:  # jeśli pole "Temperatura otoczenia" nie jest puste
                temperatura_otoczenia = float(temperatura_otoczenia_entry)
            else:  # jeśli pole "Temperatura otoczenia" jest puste
                if aktualny_miesiac in self.temperatura_miesieczna:  # jeśli podano temperaturę dla aktualnego miesiąca
                    temperatura_otoczenia = self.temperatura_miesieczna[aktualny_miesiac]
                else:  # jeśli nie podano temperatury ani w polu "Temperatura otoczenia", ani dla aktualnego miesiąca
                    self.result_label.config(text=self.texts['enter_environment_temperature'])
                    return

            # Obliczenia
            energia = sum(
                [self.oblicz_wytworzona_energie(powierzchnia, kat_nachylenia, azymut, moc_paneli, sprawność_paneli,
                                                self.temperatura_miesieczna.get(i + 1, 25),
                                                self.naswietlenie_miesieczne.get(i + 1, domyslne_naswietlenie[i + 1]),
                                                liczba_godzin_słonecznych, albedo, zacienienie, zanieczyszczenia) for i
                 in range(12)])
            self.result_label.config(text=f"{self.texts['generated_energy']}: {energia:.2f} kWh")

            # Wywołanie funkcji wykres_slupkowy
            self.wykres_slupkowy(miesiace, powierzchnia, kat_nachylenia, azymut, moc_paneli, sprawność_paneli,
                                 liczba_godzin_słonecznych, domyslne_naswietlenie, albedo, zacienienie,
                                 zanieczyszczenia)

        except ValueError as e:
            self.result_label.config(text=self.texts[str(e)])

    def wykres_slupkowy(self, miesiace, powierzchnia, kat_nachylenia, azymut, moc_paneli, sprawność_paneli,
                        liczba_godzin_słonecznych, domyslne_naswietlenie, albedo, zacienienie, zanieczyszczenia):
        # Aktualizacja wykresu
        self.ax.clear()
        bars = self.ax.bar(range(1, 13),
                           [self.oblicz_wytworzona_energie(powierzchnia, kat_nachylenia, azymut, moc_paneli,
                                                           sprawność_paneli,
                                                           self.temperatura_miesieczna.get(i + 1, 25),
                                                           self.naswietlenie_miesieczne.get(i + 1,
                                                                                            domyslne_naswietlenie[
                                                                                                i + 1]),
                                                           liczba_godzin_słonecznych, albedo, zacienienie,
                                                           zanieczyszczenia) for i in range(12)])
        self.ax.set_ylabel(self.texts['generated_energy_kWh'])
        self.ax.set_xticks(range(1, 13))
        self.ax.set_xticklabels(miesiace)

        # Dodanie etykiet do słupków
        for bar in bars:
            yval = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width() / 2.0, yval + 0.05, round(yval, 2),
                         va='bottom', ha='center')  # va: wyrównanie pionowe, ha: wyrównanie poziome

        self.canvas.draw()

    def zapisz_wykres(self):
        # Podpowiedź dla nazwy pliku
        lista_plikow = glob.glob('wykres*.png')
        numery_plikow = [int(plik.split('.')[0].split('wykres')[1]) for plik in lista_plikow if
                         plik.split('.')[0].split('wykres')[1].isdigit()]
        nowy_numer = max(numery_plikow) + 1 if numery_plikow else 1
        propozycja_nazwy = f'wykres{nowy_numer}.png'

        # Pytaj użytkownika o nazwę pliku
        nazwa_pliku = askstring(self.texts['file_name_prompt'], self.texts['enter_file_name'],
                                initialvalue=propozycja_nazwy, texts=self.texts)

        # Zapisz wykres do pliku, jeśli nazwa_pliku nie jest None
        if nazwa_pliku is not None:
            self.fig.savefig(nazwa_pliku)
            messagebox.showinfo(self.texts['info'], f"{self.texts['chart_saved']} {nazwa_pliku} {self.texts['in_this_folder']}")

    def zapisz_ustawienia(self):
        # Podpowiedź dla nazwy pliku
        lista_plikow = glob.glob('ustawienia*.xlsx')
        numery_plikow = [int(plik.split('.')[0].split('ustawienia')[1]) for plik in lista_plikow if
                         plik.split('.')[0].split('ustawienia')[1].isdigit()]
        nowy_numer = max(numery_plikow) + 1 if numery_plikow else 1
        propozycja_nazwy = f'ustawienia{nowy_numer}.xlsx'

        # Pytaj użytkownika o nazwę pliku
        nazwa_pliku = askstring(self.texts['file_name_prompt'], self.texts['enter_file_name'],
                                initialvalue=propozycja_nazwy, texts=self.texts)

        # Zapisz ustawienia do pliku, jeśli nazwa_pliku nie jest None
        if nazwa_pliku is not None:
            df = pd.DataFrame({'naswietlenie_miesieczne': list(self.naswietlenie_miesieczne.values()),
                               'temperatura_miesieczna': list(self.temperatura_miesieczna.values())},
                              index=list(self.naswietlenie_miesieczne.keys()))
            df.to_excel(nazwa_pliku)
            messagebox.showinfo(self.texts['info'],
                                f"{self.texts['settings_saved']} {nazwa_pliku} {self.texts['in_this_folder']}")

    def zapisz_dane(self):
        # Podpowiedź dla nazwy pliku
        lista_plikow = glob.glob('dane*.xlsx')
        numery_plikow = [int(plik.split('.')[0].split('dane')[1]) for plik in lista_plikow if
                         plik.split('.')[0].split('dane')[1].isdigit()]
        nowy_numer = max(numery_plikow) + 1 if numery_plikow else 1
        propozycja_nazwy = f'dane{nowy_numer}.xlsx'

        # Pytaj użytkownika o nazwę pliku
        nazwa_pliku = askstring(self.texts['file_name_prompt'], self.texts['enter_file_name'],
                                initialvalue=propozycja_nazwy, texts=self.texts)

        # Zapisz dane do pliku, jeśli nazwa_pliku nie jest None
        if nazwa_pliku is not None:
            dane = {label: entry.get() for label, entry in self.entries.items()}
            df = pd.DataFrame(dane, index=[0])
            df.to_excel(nazwa_pliku, index=False)
            messagebox.showinfo(self.texts['info'], f"{self.texts['data_saved']} {nazwa_pliku}  {self.texts['in_this_folder']}")

    def importuj_wykres(self):
        # Pozwól użytkownikowi wybrać plik do importu
        nazwa_pliku = filedialog.askopenfilename(filetypes=[("Pliki PNG", "*.png")])
        if nazwa_pliku:
            self.ax.clear()  # Czyszczenie wykresu
            img = plt.imread(nazwa_pliku)
            self.ax.imshow(img, aspect='auto')  # Ustawienie zachowania proporcji
            self.ax.axis('off')  # Wyłączenie wyświetlania osi
            self.canvas.draw()  # Aktualizacja wykresu
            messagebox.showinfo(self.texts["success"], self.texts['chart_imported'])

    def sprawdz_plik(self, df, oczekiwane_kolumny):
        # Sprawdź, czy DataFrame ma odpowiednią liczbę kolumn
        if len(df.columns) != len(oczekiwane_kolumny):
            raise ValueError(self.texts['wrong_number_of_columns'])

        # Sprawdź, czy wszystkie oczekiwane kolumny są obecne
        for kolumna in oczekiwane_kolumny:
            if kolumna not in df.columns:
                raise ValueError(f"{self.texts['missing_expected_column']}: {kolumna}")

        # Jeśli wszystko jest w porządku, zwróć True
        return True

    def wybierz_i_importuj_ustawienia(self):
        try:
            # Pozwól użytkownikowi wybrać plik do importu
            nazwa_pliku = filedialog.askopenfilename(filetypes=[(self.texts["excel_files"], "*.xlsx")])
            if nazwa_pliku:
                df = pd.read_excel(nazwa_pliku, index_col=0)

                # Sprawdź, czy plik jest poprawny
                if not self.sprawdz_plik(df, ['naswietlenie_miesieczne', 'temperatura_miesieczna']):
                    raise ValueError(self.texts["file_is_incorrect_or_damaged"])

                self.naswietlenie_miesieczne = {i + 1: v for i, v in enumerate(df['naswietlenie_miesieczne'].tolist())}
                self.temperatura_miesieczna = {i + 1: v for i, v in enumerate(df['temperatura_miesieczna'].tolist())}

                # Aktualizacja interfejsu użytkownika z nowymi wartościami
                for i, miesiac in enumerate(self.texts['months'], start=1):
                    if hasattr(self, 'naswietlenie_entries') and i in self.naswietlenie_entries and \
                            self.naswietlenie_entries[i].winfo_exists() and i in self.naswietlenie_miesieczne:
                        self.naswietlenie_entries[i].delete(0, tk.END)
                        self.naswietlenie_entries[i].insert(0, self.naswietlenie_miesieczne[i])

                    if hasattr(self, 'temperatura_entries') and i in self.temperatura_entries and \
                            self.temperatura_entries[i].winfo_exists() and i in self.temperatura_miesieczna:
                        self.temperatura_entries[i].delete(0, tk.END)
                        self.temperatura_entries[i].insert(0, self.temperatura_miesieczna[i])
                messagebox.showinfo(self.texts["success"], self.texts["settings_imported_successfully"])
        except Exception as e:
            messagebox.showerror(self.texts["error"], str(e))

    def wybierz_i_importuj_dane(self):
        try:
            # Pozwól użytkownikowi wybrać plik do importu
            nazwa_pliku = filedialog.askopenfilename(filetypes=[(self.texts["excel_files"], "*.xlsx")])
            if nazwa_pliku:
                df = pd.read_excel(nazwa_pliku)

                # Sprawdź, czy plik jest poprawny
                if not self.sprawdz_plik(df, list(self.entries.keys())):
                    raise ValueError(self.texts["file_is_incorrect_or_damaged"])

                dane = df.to_dict(orient='records')[0]

                # Aktualizacja interfejsu użytkownika z nowymi wartościami
                for label, wartosc in dane.items():
                    if label in self.entries:
                        self.entries[label].delete(0, tk.END)
                        self.entries[label].insert(0, wartosc)
                messagebox.showinfo(self.texts["success"], self.texts["data_imported_successfully"])
        except Exception as e:
            messagebox.showerror(self.texts["error"], str(e))

    def pokaz_pomoc(self, event=None):
        # Tworzenie nowego okna z informacjami o obsłudze programu
        pomoc_okno = tk.Toplevel(self.root)
        pomoc_okno.title(self.texts['help'])

        # Ustawienie stylu okna
        pomoc_okno.configure(bg='white')

        # Wczytywanie informacji z pliku Pomoc.txt
        help_file = 'Help.txt' if self.language == 'english' else 'Pomoc.txt'
        try:
            with open(os.path.join('Pomoc', help_file), 'r', encoding='utf-8') as f:
                pomoc_tekst = f.read().split('---')  # Podział tekstu na fragmenty
        except FileNotFoundError:
            pomoc_tekst = [self.texts['help_not_found']]

        # Dodanie obsługi obrazów
        images = ['obraz1.jpg', 'obraz2.png']  # Lista obrazów do wyświetlenia
        for i, text_chunk in enumerate(pomoc_tekst):
            text = tk.Text(pomoc_okno, bg='white', fg='black', padx=20, pady=20,
                           wrap=tk.WORD)  # Ustawienie zawijania tekstu na podstawie wyrazów
            text.insert(tk.END, text_chunk)
            text.pack(expand=True, fill='both')  # Zwiększenie wysokości okna z tekstem

            if i < len(images):
                image_path = os.path.join('Pomoc', images[i])
                if os.path.isfile(image_path):
                    image = Image.open(image_path)
                    photo = ImageTk.PhotoImage(image)
                    label = tk.Label(pomoc_okno, image=photo)
                    label.image = photo
                    label.pack()

        # Dostosowanie rozmiaru okna do zawartości
        pomoc_okno.update()
        pomoc_okno.geometry(
            f'{text.winfo_width() + 60}x{text.winfo_height() + 20}')  # Zwiększenie szerokości i wysokości okna

        # Zablokowanie możliwości zmiany rozmiaru okna
        pomoc_okno.resizable(False, False)

    def wyjdz(self):
        if messagebox.askokcancel(self.texts['exit'], self.texts['exit_confirmation']):
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MojaAplikacja(root)
    root.mainloop()