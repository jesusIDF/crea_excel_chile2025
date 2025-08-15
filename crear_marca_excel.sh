#!/bin/bash

python rasura_scrapp_IDF.py JMC ../../exceles/IDF_Chile_2025.xlsx ../../exceles/IDF_Chile_JMC_2025.xlsx

python copia_diccionarios.py JMC ../../exceles/IDF_Individual_Fields.xlsx ../../exceles/IDF_Chile_JMC_2025.xlsx

python crea_columnas_filtros.py JMC ../../exceles/IDF_Chile_JMC_2025.xlsx
