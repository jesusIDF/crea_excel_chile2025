#!/bin/bash

python rasura_scrapp_IDF.py Ford ../exceles/IDF_Chile_2025.xlsx ../exceles/IDF_Chile_Ford_2025.xlsx

python copia_diccionarios.py Ford ../exceles/IDF_Individual_Fields.xlsx ../exceles/IDF_Chile_Ford_2025.xlsx

python crea_columnas_filtros.py Ford ../exceles/IDF_Chile_Ford_2025.xlsx
