#!/bin/bash

python rasura_scrapp_IDF.py JAC ../../exceles/IDF_Chile_2025.xlsx ../../exceles/IDF_Chile_JAC_2025.xlsx

python copia_diccionarios.py JAC ../../exceles/IDF_Individual_Fields.xlsx ../../exceles/IDF_Chile_JAC_2025.xlsx

python crea_columnas_filtros.py JAC ../../exceles/IDF_Chile_JAC_2025.xlsx
