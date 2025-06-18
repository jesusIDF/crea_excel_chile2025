#!/bin/bash

python rasura_scrapp_IDF.py Mitsubishi ../exceles/IDF_Chile_2025.xlsx ../exceles/IDF_Chile_Mitsubishi_2025.xlsx

python copia_diccionarios.py Mitsubishi ../exceles/IDF_Individual_Fields.xlsx ../exceles/IDF_Chile_Mitsubishi_2025.xlsx

python crea_columnas_filtros.py Mitsubishi ../exceles/IDF_Chile_Mitsubishi_2025.xlsx
