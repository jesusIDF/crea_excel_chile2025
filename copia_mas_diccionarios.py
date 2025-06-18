import pandas as pd
import argparse

# Definir los argumentos
print("Copia mas diccionarios ...")
parser = argparse.ArgumentParser(description="Copiar diccionarios a partir de IDF_Individual_Fields")
parser.add_argument("marca", help="Marca de auto a preservar")
parser.add_argument("archivo_entrada", help="Ruta del archivo Excel de entrada")
parser.add_argument("archivo_salida", help="Ruta del archivo Excel de salida")

# Parsear argumentos
args = parser.parse_args()

print(f"Marca a filtrar: {args.marca}")
print(f"Archivo de entrada: {args.archivo_entrada}")
print(f"Archivo de salida: {args.archivo_salida}")

# Ruta del archivo original
archivo_entrada = args.archivo_entrada.strip()  # cámbialo por el nombre real
hojas_a_filtrar = ['V To Engine', 'V To Transmission']  # reemplaza con los nombres reales de las hojas
marca_objetivo = args.marca.strip().lower()
filtro_vtn = {"car", "truck", "van"}
# Diccionario con hojas que tienen el encabezado en la fila 2 (índice 1)
hojas_con_header_en_fila2 = {
    'V To Engine': 1  # Ajusta el nombre según corresponda
}
# Columnas vacías que se agregarán a la hoja extra
columnas_extra = [
    "liters", "VLOOKUP liters",
    "CC", "VLOOKUP CC",
    "TransmissionNumSpeeds", "VLOOKUP TransmissionNumSpeeds",
    "TransmissionControlTypeName", "VLOOKUP TransmissionControlTypeName"
]

# Ruta del archivo de salida
# archivo_salida = './exceles/IDF_Chile_Chevrolet_Complemento_2025'
archivo_salida = args.archivo_salida.strip()

# Abrir archivo y procesar las hojas
with pd.ExcelFile(archivo_entrada) as xls:
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        for nombre_hoja in hojas_a_filtrar:
            # Usar header=1 si la hoja está en el diccionario, sino header=0
            header_row = hojas_con_header_en_fila2.get(nombre_hoja, 0)

            df = pd.read_excel(xls, sheet_name=nombre_hoja, header=header_row)
            
            if 'Make' not in df.columns:
                print(f"❌ La hoja '{nombre_hoja}' no tiene columna 'Make'. Se omite.")
                continue
            col_tipo = 'V Type' if 'V Type' in df.columns else 'V Type Name' if 'V Type Name' in df.columns else None
            
            if col_tipo:
                df_filtrado = df[
                    (df['Make'].astype(str).str.strip().str.lower() == marca_objetivo.lower()) &
                    (df[col_tipo].astype(str).str.strip().str.lower().isin(t.lower() for t in filtro_vtn))
                ]
            else:
                df_filtrado = pd.DataFrame()  # vacío si no hay columna válida
            
            # Guardar hoja filtrada ej. ./exceles/IDF_Chile_Chevrolet_Complemento_2025
            df_filtrado.to_excel(writer, sheet_name=nombre_hoja, index=False)

        # Crear hoja adicional con columnas vacías
        df_complemento = pd.DataFrame(columns=columnas_extra)
        df_complemento.to_excel(writer, sheet_name='IDF_Chile_Chevrolet_2025', index=False)


print(f"✅ Hojas filtradas y guardadas en: {archivo_salida}")
