#!/bin/bash

python rasura_scrapp_IDF.py Volkswagen ../../exceles/IDF_Chile_2025.xlsx ../../exceles/IDF_Chile_Volkswagen_2025.xlsx

python copia_diccionarios.py Volkswagen ../../exceles/IDF_Individual_Fields.xlsx ../../exceles/IDF_Chile_Volkswagen_2025.xlsx

python crea_columnas_filtros.py Volkswagen ../../exceles/IDF_Chile_Volkswagen_2025.xlsx
