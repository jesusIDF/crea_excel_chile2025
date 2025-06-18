#!/bin/bash

python rasura_scrapp_IDF.py Ssangyong ../exceles/IDF_Chile_2025.xlsx ../exceles/IDF_Chile_Ssangyong_2025.xlsx

python copia_diccionarios.py Ssangyong ../exceles/IDF_Individual_Fields.xlsx ../exceles/IDF_Chile_Ssangyong_2025.xlsx

python crea_columnas_filtros.py Ssangyong ../exceles/IDF_Chile_Ssangyong_2025.xlsx
