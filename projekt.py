import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import geopandas as gpd

#Ścieżka do folderu głównego, w którym znajdują się podfoldery z plikami .xlsx
sciezka_folderu = r'C:\Users\olafo\OneDrive\Pulpit\Wiktor_projekt'

#Ścieżka do pliku shapefile z granicami krajów
shapefile_path = r'C:\Users\olafo\OneDrive\Pulpit\Wiktor_projekt\assets\ne_110m_admin_0_countries\ne_110m_admin_0_countries.shp'


#Funkcja do wczytania i przetworzenia danych z pliku Excel
def przetworz_plik(sciezka_pliku):
    #Sprawdzenie, czy plik istnieje
    if not os.path.exists(sciezka_pliku):
        print(f"Plik {sciezka_pliku} nie istnieje.")
        return None

    #Wczytanie danych z pliku, wiersz 11 (header=10) jako nagłówki
    df = pd.read_excel(sciezka_pliku, sheet_name='Sheet 1', header=10)

    #Usunięcie wszystkich kolumn zawierających 'Unnamed' w nagłówku
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    #Usunięcie wierszy i kolumn całkowicie pustych
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')

    #Sprawdzenie obecności kolumny 'TIME'
    if 'TIME' not in df.columns:
        print(f"Brak kolumny 'TIME' w pliku {sciezka_pliku}. Dostępne kolumny: {df.columns}")
        return None

    #Ustawienie 'TIME' jako indeks
    df.set_index('TIME', inplace=True)

    #Zamiana 'b', 'd', ':' i pustych komórek na NaN
    df.replace(['b', 'd', ':'], np.nan, inplace=True)

    #Konwersja danych na wartości numeryczne
    df = df.apply(pd.to_numeric, errors='coerce')

    #Wypełnienie braków średnią w danej kolumnie
    df_przetworzone = df.fillna(df.mean())

    return df_przetworzone


#Listy krajów UE i poza UE
kraje_ue = [
    'Germany', 'France', 'Italy', 'Spain', 'Netherlands', 'Poland', 'Austria',
    'Finland', 'Sweden', 'Denmark', 'Belgium', 'Czechia', 'Greece', 'Ireland',
    'Slovakia', 'Portugal', 'Romania', 'Hungary', 'Lithuania', 'Latvia',
    'Slovenia', 'Estonia', 'Luxembourg', 'Malta', 'Cyprus', 'Croatia'
]
kraje_poza_ue = [
    'Switzerland', 'Norway', 'Iceland', 'Türkiye', 'Bosnia and Herzegovina',
    'Montenegro', 'North Macedonia', 'Serbia'
]

#Ustalanie listy krajów na podstawie PKB
top_5_ue = ['Germany', 'France', 'Italy', 'Spain', 'Netherlands']
bottom_5_ue = ['Bulgaria', 'Romania', 'Latvia', 'Lithuania', 'Slovakia']

top_5_poza_ue = ['Switzerland', 'Norway', 'Iceland', 'Türkiye', 'Serbia']
bottom_5_poza_ue = ['Bosnia and Herzegovina', 'Montenegro', 'North Macedonia']

#-- Wczytanie shapefile tylko raz (jeśli istnieje) --
if not os.path.exists(shapefile_path):
    print(f"Plik shapefile nie istnieje: {shapefile_path}")
    print("Mapy nie będą generowane.")
    world = None
else:
    world = gpd.read_file(shapefile_path)
    if 'SOVEREIGNT' not in world.columns:
        print(f"Brak kolumny 'SOVEREIGNT' w shapefile. Mapy nie będą generowane.")
        world = None

#-----------------------------
#GŁÓWNA PĘTLA PO PODFOLDERACH
#-----------------------------
for folder in os.listdir(sciezka_folderu):
    sciezka_folderu_full = os.path.join(sciezka_folderu, folder)

    if os.path.isdir(sciezka_folderu_full):
        #Pętla przez pliki .xlsx w podfolderze
        for nazwa_pliku in os.listdir(sciezka_folderu_full):
            if nazwa_pliku.endswith('.xlsx'):
                sciezka_pliku = os.path.join(sciezka_folderu_full, nazwa_pliku)

                #Przetwarzanie pliku
                df_przetworzone = przetworz_plik(sciezka_pliku)

                if df_przetworzone is not None:

                    #---------------------------------------------------------
                    #1. GENEROWANIE WYKRESÓW LINIOWYCH I SŁUPKOWYCH
                    #---------------------------------------------------------

                    #Przygotowanie danych agregowanych
                    df_top_5_ue = df_przetworzone.loc[top_5_ue].mean(axis=0) if all(
                        kraj in df_przetworzone.index for kraj in top_5_ue) else df_przetworzone.loc[kraje_ue].mean(axis=0)

                    df_bottom_5_ue = df_przetworzone.loc[bottom_5_ue].mean(axis=0) if all(
                        kraj in df_przetworzone.index for kraj in bottom_5_ue) else df_przetworzone.loc[kraje_ue].mean(axis=0)

                    df_top_5_poza_ue = df_przetworzone.loc[top_5_poza_ue].mean(axis=0) if all(
                        kraj in df_przetworzone.index for kraj in top_5_poza_ue) else df_przetworzone.loc[kraje_poza_ue].mean(axis=0)

                    df_bottom_5_poza_ue = df_przetworzone.loc[bottom_5_poza_ue].mean(axis=0) if all(
                        kraj in df_przetworzone.index for kraj in bottom_5_poza_ue) else df_przetworzone.loc[kraje_poza_ue].mean(axis=0)

                    #Legenda
                    legend_labels = {
                        'Top 5 Krajów UE': top_5_ue,
                        'Bottom 5 Krajów UE': bottom_5_ue,
                        'Top 5 Krajów Poza UE': top_5_poza_ue,
                        'Bottom 5 Krajów Poza UE': bottom_5_poza_ue,
                        'Polska': ['Poland']
                    }

                    #Nazwa pliku wykresu
                    nazwa_bez_ext = os.path.splitext(nazwa_pliku)[0]

                    #---------------------
                    #WYKRES LINIOWY
                    #---------------------
                    plt.figure(figsize=(10, 6))

                    plt.plot(df_przetworzone.columns, df_top_5_ue,
                             label='Top 5 Krajów UE', color='blue', marker='o', linestyle='-', markersize=6)
                    plt.plot(df_przetworzone.columns, df_bottom_5_ue,
                             label='Bottom 5 Krajów UE', color='red', marker='x', linestyle='-', markersize=6)
                    plt.plot(df_przetworzone.columns, df_top_5_poza_ue,
                             label='Top 5 Krajów Poza UE', color='green', marker='o', linestyle='-', markersize=6)
                    plt.plot(df_przetworzone.columns, df_bottom_5_poza_ue,
                             label='Bottom 5 Krajów Poza UE', color='purple', marker='x', linestyle='-', markersize=6)

                    #Sprawdzenie, czy w indeksie jest 'Poland'
                    if 'Poland' in df_przetworzone.index:
                        plt.plot(df_przetworzone.columns, df_przetworzone.loc['Poland'],
                                 label='Polska', color='darkred', marker='s', linestyle='-', markersize=6)

                    plt.title(f'Wykres Liniowy dla {nazwa_pliku}', fontsize=14)
                    plt.xlabel('Rok', fontsize=12)
                    plt.ylabel('Procent', fontsize=12)
                    plt.grid(True)
                    plt.xticks(rotation=45, ha='right')

                    #Dodanie legendy z opisem
                    plt.legend(labels=[
                        f'{label}: {", ".join(kraje)}'
                        for label, kraje in legend_labels.items()
                    ],
                        loc='upper center',
                        bbox_to_anchor=(0.5, -0.15),
                        ncol=1)

                    plt.tight_layout()

                    #Wyświetlenie wykresu linowego
                    plt.show()

                    #---------------------
                    #WYKRES SŁUPKOWY
                    #---------------------
                    plt.figure(figsize=(12, 6))
                    plt.barh(df_przetworzone.columns, df_top_5_ue,
                             color='blue', alpha=0.7, label='Top 5 UE')
                    plt.barh(df_przetworzone.columns, df_bottom_5_ue,
                             color='red', alpha=0.7, label='Bottom 5 UE')
                    plt.barh(df_przetworzone.columns, df_top_5_poza_ue,
                             color='green', alpha=0.7, label='Top 5 Poza UE')
                    plt.barh(df_przetworzone.columns, df_bottom_5_poza_ue,
                             color='purple', alpha=0.7, label='Bottom 5 Poza UE')

                    #Polska
                    if 'Poland' in df_przetworzone.index:
                        plt.barh(df_przetworzone.columns, df_przetworzone.loc['Poland'],
                                 color='darkred', alpha=0.7, label='Polska')

                    plt.title(f'Wykres Słupkowy dla {nazwa_pliku}', fontsize=14)
                    plt.xlabel('Procent', fontsize=12)
                    plt.ylabel('Rok', fontsize=12)

                    #Dodanie legendy
                    plt.legend(labels=[
                        f'{label}: {", ".join(kraje)}'
                        for label, kraje in legend_labels.items()
                    ],
                        loc='upper center',
                        bbox_to_anchor=(0.5, -0.15),
                        ncol=1)

                    plt.tight_layout()
                    plt.show()

                    #---------------------------------------------------------
                    #2. GENEROWANIE MAPY
                    #---------------------------------------------------------
                    if world is not None:
                        #Filtrowanie krajów (UE + poza UE) w shapefile
                        europe = world[world['SOVEREIGNT'].isin(kraje_ue + kraje_poza_ue)].copy()

                        def oblicz_srednia(kraj):
                            if kraj in df_przetworzone.index:
                                return df_przetworzone.loc[kraj].mean()
                            else:
                                return np.nan

                        europe['value'] = europe['SOVEREIGNT'].apply(oblicz_srednia)

                        #Rysowanie mapy
                        fig, ax = plt.subplots(figsize=(10, 8))

                        #Obrys granic i wypełnienie kolorem
                        europe.boundary.plot(ax=ax, color='black', linewidth=0.5)
                        europe.plot(
                            column='value',
                            ax=ax,
                            legend=True,
                            legend_kwds={
                                'label': "Średnia wartość procentowa",
                                'orientation': "horizontal"
                            },
                            cmap='coolwarm'
                        )

                        #granice Europy
                        ax.set_xlim(-25, 45)
                        ax.set_ylim(35, 72)
                        plt.title(f'Mapa Europy - {nazwa_pliku}', fontsize=14)

                        plt.tight_layout()
                        plt.show()
