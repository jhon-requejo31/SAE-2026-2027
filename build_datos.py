#!/usr/bin/env python3
"""Reemplaza a excel_a_datos.py y unir_datos.py: lee el Listado (3 periodos en una sola hoja) y genera datos.js compacto.
Uso: python build_datos.py Listado_Servicios_Educativos_PC_2026_2027.xlsx datos.js"""
import sys, json, pandas as pd
KEEP = """periodo_sae COD_MODULAR ANEXO COD_LOCAL UNIDAD_TERRITORIAL DEPARTAMENTO PROVINCIA DISTRITO NOMBRE_IE NIVEL USUARIOS COMITÉ ITEM modalidad forma_atencion
TIPO_RACION REGION_ALIMENTARIA AREA Quintil grupo_sae_educunas estado_ie Latitud Longitud rango_usuarios iieeconfichadeinfra2No25 iieedetalle kitmenaje kitcocina kitmobiliario
ambientecocina ambientealmacen ambientecomedor abastecimiento_agua_por servicio_higienico_conectado_a alumbrado_proviene_de material_predominante_piso_almac
material_predominante_piso_cocin material_predominante_piso_comed material_predominante_paredes_al material_predominante_paredes_co material_predominante_pared_com
material_predominante_techo_alma material_predominante_techo_coci material_predominante_techo_come cuenta_cocina_adecuada agua_adecuada desague_adecuado
electricidad_adecuada menaje_buen_estado kitcocina_buen_estado kitmobiliario_buen_estado cocina_optima""".split()
A = ''.join(chr(i) for i in range(35, 127) if i != 92)
df = pd.read_excel(sys.argv[1], dtype=str).fillna('')[KEEP].apply(lambda c: c.str.strip())
df.columns = [c.lower().replace('é', 'e') for c in df.columns]
df['unidad_territorial'] = df['unidad_territorial'].str.upper()
out = {'n': len(df), 'A': A, 'str': {}, 'int': {}, 'num': {}, 'dict': {}}
for c in df.columns:
    if c in ('usuarios', 'latitud', 'longitud'):
        v = pd.to_numeric(df[c].str.replace(',', '.'), errors='coerce').fillna(0)
        out['num'][c] = [round(x, 5) if c != 'usuarios' else int(x) for x in v]
        continue
    cat = pd.Categorical(df[c]); out['dict'][c] = [str(x) for x in cat.categories]
    if len(cat.categories) <= len(A): out['str'][c] = ''.join(A[k] for k in cat.codes)
    else: out['int'][c] = cat.codes.tolist()
open(sys.argv[2], 'w', encoding='utf8').write('window.SAE=' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + ';')
print('Listo', sys.argv[2], len(df), 'filas')
