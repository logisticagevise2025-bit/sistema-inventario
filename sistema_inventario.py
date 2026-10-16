import openpyxl
import re
import os
from copy import copy

# 1. ENTORNO DE DIRECTORIOS LOCALES
carpeta_actual = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()

# Nombres de archivos originales estrictos solicitados por el usuario
archivo_fuente = os.path.join(carpeta_actual, 'INVENTARIO PUESTOS GEVISE 2026.xlsx')
archivo_plantilla = os.path.join(carpeta_actual, 'FORMATO DE CONTROL DE INVENTARIO GEVISE.xlsx')

if not os.path.exists(archivo_fuente):
    raise FileNotFoundError(f"Error crítico: No se encuentra la fuente de verdad '{os.path.basename(archivo_fuente)}' en: {carpeta_actual}")
if not os.path.exists(archivo_plantilla):
    raise FileNotFoundError(f"Error crítico: No se encuentra la plantilla base '{os.path.basename(archivo_plantilla)}' en: {carpeta_actual}")

# Cargar fuentes de datos principales
wb_fuente_maestra = openpyxl.load_workbook(archivo_fuente, data_only=True)
hoja_maestra = wb_fuente_maestra.active

# 2. MAPEO ROBUSTO DINÁMICO POR ENCABEZADOS (FILA 1)
columnas_map = {}
for col_idx in range(1, hoja_maestra.max_column + 1):
    val_header = hoja_maestra.cell(row=1, column=col_idx).value
    if val_header:
        columnas_map[str(val_header).strip().upper()] = col_idx

# Definición de correspondencias según los requerimientos del usuario
map_directo_celdas = {
    "C12": columnas_map.get("FORRO DE CHALECO"),
    "C13": columnas_map.get("PLACAS ANTIBALAS"),
    "D13": columnas_map.get("SERIE"),
    "C14": columnas_map.get("CINTO"),
    "C15": columnas_map.get("PORTA ARMA"),
    "C16": columnas_map.get("PORTA RADIO"),
    "C17": columnas_map.get("PORTA TOLETE"),
    "C18": columnas_map.get("TOLETE"),
    "C21": columnas_map.get("CELULAR COPORATIVO"),
    "C22": columnas_map.get("LINTERNA"),
    "C23": columnas_map.get("PONCHO IMPERMEABLE"),
    "C24": columnas_map.get("BOTAS IMPERMEABLES"),
    "C25": columnas_map.get("CHALECO REFLECTIVO"),
    "C26": columnas_map.get("CASCO"),
    "C27": columnas_map.get("DETECTOR METALES"),
    "C29": columnas_map.get("MESA"),
    "C30": columnas_map.get("SILLA"),
    "C31": columnas_map.get("CARPA"),
    "C32": columnas_map.get("GARITA"),
    "C33": columnas_map.get("CONOS TRANSITO"),
    "C34": columnas_map.get("CAJA FUERTE"),
    "C35": columnas_map.get("CAJA METALICA"),
    "C36": columnas_map.get("CANDADO")
}

col_proyecto_b = columnas_map.get("NOMBRE DEL PUESTO")
col_nominativo_c = columnas_map.get("NOMINATIVO")
col_armamento_l = columnas_map.get("ARMAMENTO")
col_municion_m = columnas_map.get("MUNICION")
col_radio_n = columnas_map.get("RADIO PORTATIL")
col_adicionales_ad = columnas_map.get("ADICIONALES")

def procesar_articulo_adicional(texto_adicional):
    texto_adicional = texto_adicional.strip()
    match = re.match(r"^(\d+)\s+(.*)$", texto_adicional)
    if match:
        return int(match.group(1)), match.group(2).strip().upper()
    return 1, texto_adicional.upper()

# Cargar molde limpio de plantilla para extraer la estructura de firmas original (Filas >= 40)
wb_molde_firmas = openpyxl.load_workbook(archivo_plantilla)
hoja_molde_firmas = wb_molde_firmas.active
lista_combinadas_firmas_origen = []
for rng in hoja_molde_firmas.merged_cells.ranges:
    if rng.min_row >= 40:
        lista_combinadas_firmas_origen.append({
            "min_col": rng.min_col, "max_col": rng.max_col,
            "min_row": rng.min_row, "max_row": rng.max_row
        })
wb_molde_firmas.close()

# 3. CONSOLIDACIÓN INTERNA DE FILAS AGRUPADAS POR PROYECTO
diccionario_proyectos = {}
for f_idx in range(2, hoja_maestra.max_row + 1):
    proy_nombre = hoja_maestra.cell(row=f_idx, column=col_proyecto_b).value
    if proy_nombre and str(proy_nombre).strip():
        proy_limpio = str(proy_nombre).strip().upper()
        if proy_limpio not in diccionario_proyectos:
            diccionario_proyectos[proy_limpio] = []
        diccionario_proyectos[proy_limpio].append(f_idx)

# 4. GENERACIÓN DE ARCHIVOS INDEPENDIENTES POR PROYECTO
for proyecto_k, filas_asociadas in diccionario_proyectos.items():
    wb_salida_proyecto = openpyxl.load_workbook(archivo_plantilla)
    hoja_molde_proyecto = wb_salida_proyecto.active
    
    nombre_archivo_salida = f"INVENTARIOS {proyecto_k}.xlsx"
    ruta_salida_final = os.path.join(carpeta_actual, nombre_archivo_salida)
    
    puestos_procesados_en_proyecto = 0
    
    for fila_data_idx in filas_asociadas:
        val_proyecto_b = hoja_maestra.cell(row=fila_data_idx, column=col_proyecto_b).value
        val_nominativo_c = hoja_maestra.cell(row=fila_data_idx, column=col_nominativo_c).value
        
        sigla_pestaña = str(val_nominativo_c).strip().replace(" ", "_") if val_nominativo_c else f"PUESTO_{fila_data_idx}"
        nombre_pestaña_unico = f"{sigla_pestaña}_{fila_data_idx}"
        
        nueva_hoja_puesto = wb_salida_proyecto.copy_worksheet(hoja_molde_proyecto)
        nueva_hoja_puesto.title = nombre_pestaña_unico[0:30]
        puestos_procesados_en_proyecto += 1
        
        # Inyección de Encabezados Principales
        nueva_hoja_puesto["B3"] = val_proyecto_b
        nueva_hoja_puesto["F3"] = val_nominativo_c
        
        # Inyección de Armamento, Munición y Comunicaciones (L, M, N)
        nueva_hoja_puesto["E5"] = str(hoja_maestra.cell(row=fila_data_idx, column=col_armamento_l).value).strip() if hoja_maestra.cell(row=fila_data_idx, column=col_armamento_l).value else ""
        nueva_hoja_puesto["C6"] = hoja_maestra.cell(row=fila_data_idx, column=col_municion_m).value if hoja_maestra.cell(row=fila_data_idx, column=col_municion_m).value is not None else 0
        nueva_hoja_puesto["E8"] = str(hoja_maestra.cell(row=fila_data_idx, column=col_radio_n).value).strip() if hoja_maestra.cell(row=fila_data_idx, column=col_radio_n).value else ""
        
        # Inyección Automatizada de Filas 12 a 36 utilizando el mapeo dinámico
        for celda_dest, col_orig_idx in map_directo_celdas.items():
            if col_orig_idx:
                val_celda_origen = hoja_maestra.cell(row=fila_data_idx, column=col_orig_idx).value
                if val_celda_origen is not None:
                    nueva_hoja_puesto[celda_dest] = val_celda_origen
                else:
                    nueva_hoja_puesto[celda_dest] = "-"
            else:
                nueva_hoja_puesto[celda_dest] = "-"
        
        # --- PROCESAMIENTO DINÁMICA DE ARTÍCULOS ADICIONALES (COL AD) ---
        raw_adicionales = hoja_maestra.cell(row=fila_data_idx, column=col_adicionales_ad).value
        str_adicionales = str(raw_adicionales).strip() if raw_adicionales is not None else ""
        lista_items_adicionales = [i.strip() for i in str_adicionales.split(",") if i.strip()]
        total_items_adicionales = len(lista_items_adicionales)
        
        desfase_filas = 0
        if total_items_adicionales > 3:
            desfase_filas = total_items_adicionales - 3
            
            for block in list(nueva_hoja_puesto.merged_cells.ranges):
                if block.min_row >= 43:
                    nueva_hoja_puesto.unmerge_cells(start_row=block.min_row, start_column=block.min_col, end_row=block.max_row, end_column=block.max_col)
            
            alturas_firmas_originales = {}
            for r_firmas in range(43, 65):
                alturas_firmas_originales[r_firmas] = nueva_hoja_puesto.row_dimensions[r_firmas].height
            
            nueva_hoja_puesto.insert_rows(idx=43, amount=desfase_filas)
            
            altura_base_molde = nueva_hoja_puesto.row_dimensions.height if 42 in nueva_hoja_puesto.row_dimensions else 20
            for f_nueva in range(43, 43 + desfase_filas):
                nueva_hoja_puesto.row_dimensions[f_nueva].height = altura_base_molde
                for c_idx in range(1, 12):
                    cell_origen = nueva_hoja_puesto.cell(row=42, column=c_idx)
                    cell_destino = nueva_hoja_puesto.sw_cell if hasattr(nueva_hoja_puesto, 'sw_cell') else nueva_hoja_puesto.cell(row=f_nueva, column=c_idx)
                    if cell_origen.has_style:
                        cell_destino.font = copy(cell_origen.font)
                        cell_destino.border = copy(cell_origen.border)
                        cell_destino.fill = copy(cell_origen.fill)
                        cell_destino.alignment = copy(cell_origen.alignment)
            
            for r_orig, alt_val in alturas_firmas_originales.items():
                if alt_val is not None:
                    nueva_hoja_puesto.row_dimensions[r_orig + desfase_filas].height = alt_val
            
            for comb in lista_combinadas_firmas_origen:
                if comb["min_row"] >= 43:
                    nueva_hoja_puesto.merge_cells(
                        start_row=comb["min_row"] + desfase_filas, start_column=comb["min_col"],
                        end_row=comb["max_row"] + desfase_filas, end_column=comb["max_col"]
                    )
        
        limite_filas_adicionales = 40 + max(3, total_items_adicionales)
        rango_filas_totales_adicionales = list(range(40, limite_filas_adicionales))
        
        for idx_ad, fila_actual_ad in enumerate(rango_filas_totales_adicionales):
            nueva_hoja_puesto.merge_cells(start_row=fila_actual_ad, start_column=4, end_row=fila_actual_ad, end_column=11)
            for col_clean in range(4, 12):
                nueva_hoja_puesto.cell(row=fila_actual_ad, column=col_clean).value = None
                
            if idx_ad < total_items_adicionales:
                cantidad_item, nombre_item = procesar_articulo_adicional(lista_items_adicionales[idx_ad])
                nueva_hoja_puesto[f"A{fila_actual_ad}"] = nombre_item
                nueva_hoja_puesto[f"C{fila_actual_ad}"] = cantidad_item
            else:
                nueva_hoja_puesto[f"A{fila_actual_ad}"] = "OTROS: "
                nueva_hoja_puesto[f"C{fila_actual_ad}"] = None
                
    if puestos_procesados_en_proyecto > 0:
        wb_salida_proyecto.remove(hoja_molde_proyecto)
        wb_salida_proyecto.save(ruta_salida_final)
        print(f"Archivo generado exitosamente: {nombre_archivo_salida} con {puestos_procesados_en_proyecto} puestos.")
        wb_salida_proyecto.close()
        print("\n--- PROCESO GLOBAL FINALIZADO CON ÉXITO ---")