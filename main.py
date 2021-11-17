"""
Do poprawnego działania potrzebna biblioteka XlsxWriter.
W celu uruchomienia należy wskazać folder zawierający pliki w formacie '.xlsx'
"""
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# Funkcja wczytująca dane do jednej listy z wszystkich plików wsadowych.
def input_data():
    row_list_all = []
    rows_number = 0

    for tabelka in pliki_wsadowe:
        dane = pd.read_excel(tabelka)
        rows_number = len(dane.index)
        row_list_all.append(dane.to_numpy().tolist())

    return rows_number, row_list_all

# Funkcja sortująca i na bieżąco zapisująca do jednego pliku wyniki.
def output_to_wyniki():
    write_list = []
    odwierty_lista = []

    writer = pd.ExcelWriter('wyniki.xlsx', engine='xlsxwriter')
    # , options={'strings_to_formulas': False})

    for row in range(rows_number):
        for i in row_list_all:
            for j in i:
                write_list.append(j)
            odwierty_lista.append(write_list[row])
            write_list.clear()
        # print("\nPosortowane:")
        # print(odwierty_lista)

        # Ustawienie nazwy zakładki jako pierwszej komórki wiersza
        nazwa_zakladki = odwierty_lista[0][0]
        # Utworzenie DataFrame'u w celu zapisania w pliku excel
        df_odw = pd.DataFrame(odwierty_lista)
        df_odw.to_excel(writer, sheet_name=nazwa_zakladki)
        odwierty_lista.clear()
    writer.save()


if __name__ == '__main__':
    print("-" * 100)

    root = tk.Tk()
    root.withdraw()

    # Okno do wskazania folderu z plikami wsadowymi
    directory_path = filedialog.askdirectory(title="Wybierz folder z plikami wsadowymi")
    
    # Zmienia CWD na folder z plikami wsadowymi
    try:
        os.chdir(directory_path)
    except FileNotFoundError:
        print("Brak folderu!")
        exit(100)
    except OSError:
        print("Brak folderu!")
        exit(200)

    # Tworzy listę z nazwami plików wsadowych
    pliki_wsadowe = [name for name in os.listdir('.') if os.path.isfile(name)]
    if pliki_wsadowe == []:
        print("Wskazany folder jest pusty!")
        exit(300)

    rows_number, row_list_all = input_data()

    # print("Lista wszystkich wierszy: \n", row_list)
    print("Liczba wierszy w pliku: ", rows_number)
    print("-" * 100)

    output_to_wyniki()

    print("Pomyślnie utworzono plik 'wyniki.xlsx'.")
    print("-" * 100)
