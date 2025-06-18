#!/bin/bash

python rasura_scrapp_IDF.py MG ../exceles/IDF_Chile_2025.xlsx ../exceles/IDF_Chile_MG_2025.xlsx

python copia_diccionarios.py MG ../exceles/IDF_Individual_Fields.xlsx ../exceles/IDF_Chile_MG_2025.xlsx

python crea_columnas_filtros.py MG ../exceles/IDF_Chile_MG_2025.xlsx
