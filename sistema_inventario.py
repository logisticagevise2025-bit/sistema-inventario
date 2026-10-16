import openpyxl
import re
import os
import unicodedata
from copy import copy

carpeta_actual = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
archivo_fuente = os.path.join(carpeta_actual, 'INVENTARIO PUESTOS GEVISE 2026.xlsx')
archivo_plantilla = os.path.join(carpeta_actual, 'FORMATO DE CONTROL DE INVENTARIO GEVISE.xlsx')

if not os.path.exists(archivo_fuente) or not os.path.exists(archivo_plantilla):
    raise FileNotFoundError("Error: Asegúrate de tener la fuente de verdad y la plantilla en la carpeta.")

wb_fuente_maestra = openpyxl.load_workbook(archivo_fuente, data_only=True)

# 1. DEFINICIÓN ACTUALIZADA DE HOJAS VÁLIDAS POR REGIÓN
hojas_validas_regiones = ["COSTA", "SIERRA", "AMAZONIA"]


def normalizar_texto(texto):
    if not texto:
        return ""
    texto_norm = ''.join(c for c in unicodedata.normalize('NFD', str(texto).strip().upper()) if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', texto_norm)


def procesar_articulo_adicional(texto_adicional):
    match = re.match(r"^(\d+)\s+(.*)$", str(texto_adicional).strip())
    return (int(match.group(1)), match.group(2).strip().upper()) if match else (1, str(texto_adicional).strip().upper())


# Mapear encabezados usando la primera hoja válida que exista en el archivo
hoja_guia = None
for h_name in wb_fuente_maestra.sheetnames:
    if h_name.strip().upper() in hojas_validas_regiones:
        hoja_guia = wb_fuente_maestra[h_name]
        break
if hoja_guia is None:
    hoja_guia = wb_fuente_maestra.active

columnas_map = {normalizar_texto(hoja_guia.cell(row=1, column=c).value): c for c in range(1, hoja_guia.max_column + 1) if hoja_guia.cell(row=1, column=c).value}

col_forro = columnas_map.get(normalizar_texto("FORRO DE CHALECO")) or columnas_map.get(normalizar_texto("FORRO CHALECO"))
if not col_forro:
    for k, v in columnas_map.items():
        if "FORRO" in k:
            col_forro = v
            break

map_directo_celdas = {
    "C12": col_forro, "C13": columnas_map.get(normalizar_texto("PLACAS ANTIBALAS")),
    "D13": columnas_map.get(normalizar_texto("SERIE")), "C14": columnas_map.get(normalizar_texto("CINTO")),
    "C15": columnas_map.get(normalizar_texto("PORTA ARMA")), "C16": columnas_map.get(normalizar_texto("PORTA RADIO")),
    "C17": columnas_map.get(normalizar_texto("PORTA TOLETE")), "C18": columnas_map.get(normalizar_texto("TOLETE")),
    "C21": columnas_map.get(normalizar_texto("CELULAR COPORATIVO")), "C22": columnas_map.get(normalizar_texto("LINTERNA")),
    "C23": columnas_map.get(normalizar_texto("PONCHO IMPERMEABLE")), "C24": columnas_map.get(normalizar_texto("BOTAS IMPERMEABLES")),
    "C25": columnas_map.get(normalizar_texto("CHALECO REFLECTIVO")), "C26": columnas_map.get(normalizar_texto("CASCO")),
    "C27": columnas_map.get(normalizar_texto("DETECTOR METALES")), "C29": columnas_map.get(normalizar_texto("MESA")),
    "C30": columnas_map.get(normalizar_texto("SILLA")), "C31": columnas_map.get(normalizar_texto("CARPA")),
    "C32": columnas_map.get(normalizar_texto("GARITA")), "C33": columnas_map.get(normalizar_texto("CONOS TRANSITO")),
    "C34": columnas_map.get(normalizar_texto("CAJA FUERTE")), "C35": columnas_map.get(normalizar_texto("CAJA METALICA")),
    "C36": columnas_map.get(normalizar_texto("CANDADO"))
}

idx_proyecto = columnas_map.get(normalizar_texto("NOMBRE DEL PUESTO")) or 1
idx_nominativo = columnas_map.get(normalizar_texto("NOMINATIVO")) or 2
idx_armamento = columnas_map.get(normalizar_texto("ARMAMENTO")) or 11
idx_municion = columnas_map.get(normalizar_texto("MUNICION")) or 12
idx_radio = columnas_map.get(normalizar_texto("RADIO PORTATIL")) or 13
col_adicionales_ad = columnas_map.get(normalizar_texto("ADICIONALES"))

wb_molde_firmas = openpyxl.load_workbook(archivo_plantilla)
hoja_molde_firmas = wb_molde_firmas.active
lista_combinadas_firmas_origen = [{"min_col": r.min_col, "max_col": r.max_col, "min_row": r.min_row, "max_row": r.max_row} for r in hoja_molde_firmas.merged_cells.ranges if r.min_row >= 40]
wb_molde_firmas.close()

diccionario_proyectos = {}
for nombre_hoja in wb_fuente_maestra.sheetnames:
    nombre_hoja_up = nombre_hoja.strip().upper()
    if nombre_hoja_up not in hojas_validas_regiones:
        continue
        
    hoja_provincia_actual = wb_fuente_maestra[nombre_hoja]
    
    for f_idx in range(2, hoja_provincia_actual.max_row + 1):
        proy_nombre = hoja_provincia_actual.cell(row=f_idx, column=idx_proyecto).value
        nominativo_val = hoja_provincia_actual.cell(row=f_idx, column=idx_nominativo).value
        if not proy_nombre and not nominativo_val:
            continue
            
        proy_limpio = str(proy_nombre).strip().upper() if proy_nombre and str(proy_nombre).strip() else "SIN PROYECTO"
        proy_limpio = re.sub(r'[\\/*?:\[\]]', '', proy_limpio)
        
        if proy_limpio not in diccionario_proyectos:
            diccionario_proyectos[proy_limpio] = []
            
        diccionario_proyectos[proy_limpio].append({"hoja": hoja_provincia_actual, "fila": f_idx})
for proyecto_k, filas_asociadas in diccionario_proyectos.items():
    wb_salida_proyecto = openpyxl.load_workbook(archivo_plantilla)
    hoja_molde_proyecto = wb_salida_proyecto.active
    nombre_archivo_salida = f"INVENTARIOS {proyecto_k}.xlsx"
    ruta_salida_final = os.path.join(carpeta_actual, nombre_archivo_salida)
    puestos_procesados_en_proyecto = 0

    for idx_conteo, meta_data in enumerate(filas_asociadas, start=1):
        hoja_origen_física = meta_data["hoja"]
        fila_data_idx = meta_data["fila"]
        
        val_proyecto_b = hoja_origen_física.cell(row=fila_data_idx, column=idx_proyecto).value or ""
        val_nominativo_c = hoja_origen_física.cell(row=fila_data_idx, column=idx_nominativo).value or ""
        
        sigla_pestaña = str(val_nominativo_c).strip().replace(" ", "_") if val_nominativo_c else "PUESTO"
        sigla_pestaña = re.sub(r'[\\/*?:\[\]]', '', sigla_pestaña)
        nombre_pestaña_unico = f"{idx_conteo}_{sigla_pestaña}"[0:30]
        
        nueva_hoja_puesto = wb_salida_proyecto.copy_worksheet(hoja_molde_proyecto)
        nueva_hoja_puesto.title = nombre_pestaña_unico
        puestos_procesados_en_proyecto += 1
        
        nueva_hoja_puesto["B3"] = val_proyecto_b
        nueva_hoja_puesto["F3"] = val_nominativo_c
        nueva_hoja_puesto["E5"] = str(hoja_origen_física.cell(row=fila_data_idx, column=idx_armamento).value or "").strip()
        
        val_mun = hoja_origen_física.cell(row=fila_data_idx, column=idx_municion).value
        nueva_hoja_puesto["C6"] = val_mun if val_mun is not None else 0
        nueva_hoja_puesto["E8"] = str(hoja_origen_física.cell(row=fila_data_idx, column=idx_radio).value or "").strip()

        for celda_dest, col_orig_idx in map_directo_celdas.items():
            val_origen = hoja_origen_física.cell(row=fila_data_idx, column=col_orig_idx).value if col_orig_idx else None
            nueva_hoja_puesto[celda_dest] = val_origen if val_origen is not None else "-"

        raw_adicionales = hoja_origen_física.cell(row=fila_data_idx, column=col_adicionales_ad).value if col_adicionales_ad else None
        lista_items_adicionales = [i.strip() for i in str(raw_adicionales or "").split(",") if i.strip()]
        total_items_adicionales = len(lista_items_adicionales)

        if total_items_adicionales > 3:
            desfase = total_items_adicionales - 3
            for rng in list(nueva_hoja_puesto.merged_cells.ranges):
                if rng.min_row >= 43:
                    nueva_hoja_puesto.unmerge_cells(start_row=rng.min_row, start_column=rng.min_col, end_row=rng.max_row, end_column=rng.max_col)
            
            alturas = {r: nueva_hoja_puesto.row_dimensions[r].height for r in range(43, 65)}
            nueva_hoja_puesto.insert_rows(idx=43, amount=desfase)
            
            alt_base = nueva_hoja_puesto.row_dimensions.height or 20
            for f_nuevas in range(43, 43 + desfase):
                nueva_hoja_puesto.row_dimensions[f_nuevas].height = alt_base
                for c_idx in range(1, 12):
                    cell_origen = nueva_hoja_puesto.cell(row=42, column=c_idx)
                    cell_destino = nueva_hoja_puesto.cell(row=f_nuevas, column=c_idx)
                    if cell_origen.has_style:
                        cell_destino.font = copy(cell_origen.font)
                        cell_destino.border = copy(cell_origen.border)
                        cell_destino.fill = copy(cell_origen.fill)
                        cell_destino.alignment = copy(cell_origen.alignment)

            for r_orig, alt_val in alturas.items():
                if alt_val is not None:
                    nueva_hoja_puesto.row_dimensions[r_orig + desfase].height = alt_val

            for comb in lista_combinadas_firmas_origen:
                if comb["min_row"] >= 43:
                    nueva_hoja_puesto.merge_cells(start_row=comb["min_row"] + desfase, start_column=comb["min_col"], end_row=comb["max_row"] + desfase, end_column=comb["max_col"])

        rango_filas_totales_adicionales = list(range(40, 40 + max(3, total_items_adicionales)))

        for fila_actual_ad in rango_filas_totales_adicionales:
            for block in list(nueva_hoja_puesto.merged_cells.ranges):
                if block.min_row == fila_actual_ad and block.max_row == fila_actual_ad and (block.min_col >= 4 or block.max_col <= 11):
                    nueva_hoja_puesto.unmerge_cells(start_row=block.min_row, start_column=block.min_col, end_row=block.max_row, end_column=block.max_col)

            for col_clean in range(4, 12):
                nueva_hoja_puesto.cell(row=fila_actual_ad, column=col_clean).value = None

            idx_ad = fila_actual_ad - 40
            if idx_ad < total_items_adicionales:
                cantidad_item, nombre_item = procesar_articulo_adicional(lista_items_adicionales[idx_ad])
                nueva_hoja_puesto.cell(row=fila_actual_ad, column=1).value = nombre_item
                nueva_hoja_puesto.cell(row=fila_actual_ad, column=3).value = cantidad_item
            else:
                nueva_hoja_puesto.cell(row=fila_actual_ad, column=1).value = "OTROS: "
                nueva_hoja_puesto.cell(row=fila_actual_ad, column=3).value = None

            nueva_hoja_puesto.merge_cells(start_row=fila_actual_ad, start_column=4, end_row=fila_actual_ad, end_column=11)

    if puestos_procesados_en_proyecto > 0:
        wb_salida_proyecto.remove(hoja_molde_proyecto)
        wb_salida_proyecto.save(ruta_salida_final)
        print(f"Archivo generado exitosamente: {nombre_archivo_salida} con {puestos_procesados_en_proyecto} puestos consolidando regiones.")
    wb_salida_proyecto.close()

print("\n--- PROCESO GLOBAL FINALIZADO CON ÉXITO ---")
