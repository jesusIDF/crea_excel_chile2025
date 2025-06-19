#!/bin/bash

python rasura_scrapp_IDF.py Changan ../../exceles/IDF_Chile_2025.xlsx ../../exceles/IDF_Chile_Changan_2025.xlsx

python copia_diccionarios.py Changan ../../exceles/IDF_Individual_Fields.xlsx ../../exceles/IDF_Chile_Changan_2025.xlsx

python crea_columnas_filtros.py Changan ../../exceles/IDF_Chile_Changan_2025.xlsx
