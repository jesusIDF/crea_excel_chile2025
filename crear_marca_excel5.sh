#!/bin/bash

python rasura_scrapp_IDF.py Mazda ../../exceles/IDF_Chile_2025.xlsx ../../exceles/IDF_Chile_Mazda_2025.xlsx

python copia_diccionarios.py Mazda ../../exceles/IDF_Individual_Fields.xlsx ../../exceles/IDF_Chile_Mazda_2025.xlsx

python crea_columnas_filtros.py Mazda ../../exceles/IDF_Chile_Mazda_2025.xlsx
