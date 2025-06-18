#!/bin/bash

python rasura_scrapp_IDF.py Daewoo ../../exceles/IDF_Chile_2025.xlsx ../../exceles/IDF_Chile_Daewoo_2025.xlsx

python copia_diccionarios.py Daewoo ../../exceles/IDF_Individual_Fields.xlsx ../../exceles/IDF_Chile_Daewoo_2025.xlsx

python crea_columnas_filtros.py Daewoo ../../exceles/IDF_Chile_Daewoo_2025.xlsx
