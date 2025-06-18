#!/bin/bash

python rasura_scrapp_IDF.py Samsung ../exceles/IDF_Chile_2025.xlsx ../exceles/IDF_Chile_Samsung_2025.xlsx

python copia_diccionarios.py Samsung ../exceles/IDF_Individual_Fields.xlsx ../exceles/IDF_Chile_Samsung_2025.xlsx

python crea_columnas_filtros.py Samsung ../exceles/IDF_Chile_Samsung_2025.xlsx
