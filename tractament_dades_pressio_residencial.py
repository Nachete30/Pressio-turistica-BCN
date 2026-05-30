"""
tractament_dades_pressio_residencial.py
========================================
Tractament de dades per a la pràctica de visualització:
"Pressió turística i accessibilitat residencial als barris de Barcelona"

Autor: Nach (UOC)
Data: 2026-05-30

Fonts de dades:
  - Habitatge Gencat: trimestral_bcn_contractes.xlsx, trimestral_bcn_lloguer-2.xlsx,
                      trimestral_bcn_lloguer_m2.xlsx, trimestral_bcn_sup.xlsx
  - Open Data BCN:    2025_4T_hut_comunicacio_opendata_habitatges_turistics.csv
                      2022_renda_disponible_llars_per_persona.csv
                      2025_pad_mdbas.csv
                      BarcelonaCiutat_Barris.csv
  - INE (opcional):   download.do_IPC.xls (no s'utilitza; limitació documentada)

Sortides:
  - taula_final_barris_trimestre.csv
  - taula_final_barris_ultim_periode.csv
  - resum_qualitat_dades.csv
"""

import os
import re
import warnings
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# 0. CONFIGURACIÓ DE RUTES
# ─────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
GENCAT = os.path.join(BASE, "Habitatge gencat")
OPENDATA = os.path.join(BASE, "Open Data BCN")

CONTRACTES_PATH  = os.path.join(GENCAT, "trimestral_bcn_contractes.xlsx")
LLOGUER_PATH     = os.path.join(GENCAT, "trimestral_bcn_lloguer-2.xlsx")
LLOGUER_M2_PATH  = os.path.join(GENCAT, "trimestral_bcn_lloguer_m2.xlsx")
SUP_PATH         = os.path.join(GENCAT, "trimestral_bcn_sup.xlsx")
HUT_PATH         = os.path.join(OPENDATA, "2025_4T_hut_comunicacio_opendata_habitatges_turistics.csv")
RENDA_PATH       = os.path.join(OPENDATA, "2022_renda_disponible_llars_per_persona.csv")
PADRO_PATH       = os.path.join(OPENDATA, "2025_pad_mdbas.csv")
BARRIS_PATH      = os.path.join(OPENDATA, "BarcelonaCiutat_Barris.csv")

# ─────────────────────────────────────────────
# 1. FUNCIÓ AUXILIAR: NORMALITZAR NOMS DE BARRIS
# ─────────────────────────────────────────────
def normalitza_nom(s):
    """
    Neteja un nom de barri per facilitar la unió entre fonts:
    - Elimina espais en blanc als extrems i múltiples espais interns
    - Converteix a minúscules
    """
    if pd.isna(s):
        return s
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)  # espais múltiples → un sol espai
    s = s.lower()
    return s


# ─────────────────────────────────────────────
# 2. FUNCIÓ AUXILIAR: PARSEJAR UN FULL DELS XLSX DE GENCAT
# ─────────────────────────────────────────────
def parseja_full_gencat(path, any_str, metric_col_name):
    """
    Llegeix un full de qualsevol dels quatre XLSXs de Gencat.

    Estructura dels fulls (anys 2014+, ~98-100 files):
      Fila 5:   any + "any (acumulat)"  → detectem les columnes T1-T4
      Fila 6:   "Codi" | labels trimestres ("I","II","III","IV")
      Fila 7:   Barcelona total (ignorem per a la taula de barris)
      Fila 9:   secció "Districtes municipals"
      Files 10-19: 10 districtes
      Fila 21:  secció "Barris"
      Files 22-94: fins a 73 barris

    Anys anteriors a 2014 no tenen desagregació per barri → retornem buit.

    Retorna un DataFrame tidy amb columnes:
      any, trimestre, ambit, codi, nom, <metric_col_name>
    """
    wb = pd.read_excel(path, sheet_name=any_str, header=None)

    # Detectem la fila de capçalera (on hi ha "I","II","III","IV" com a etiquetes)
    header_row = None
    for i, row in wb.iterrows():
        trimestre_vals = [str(v).strip() for v in row if str(v).strip() in ['I','II','III','IV']]
        if len(trimestre_vals) >= 3:
            header_row = i
            break

    if header_row is None:
        return pd.DataFrame()

    header = wb.iloc[header_row]

    # Identifiquem les columnes dels quatre trimestres (primera aparició de cada etiqueta)
    tri_cols = {}
    for col_idx, val in enumerate(header):
        v = str(val).strip()
        if v in ['I', 'II', 'III', 'IV'] and v not in tri_cols:
            tri_cols[v] = col_idx

    if len(tri_cols) < 1:
        return pd.DataFrame()

    col_codi = 0
    col_nom  = 1

    records = []
    ambit_actual = None  # 'barcelon', 'districte' o 'barri'

    for i in range(header_row + 1, len(wb)):
        row = wb.iloc[i]
        nom_val  = row.iloc[col_nom]
        codi_val = row.iloc[col_codi]

        # Detectem línies de secció
        if pd.notna(nom_val) and pd.isna(codi_val):
            nom_str = str(nom_val).strip().lower()
            if 'districte' in nom_str:
                ambit_actual = 'districte'
                continue
            elif 'barri' in nom_str:
                ambit_actual = 'barri'
                continue
            elif 'barcelona' in nom_str:
                ambit_actual = 'barcelona'
                continue
            else:
                break  # nota al peu → aturem

        if ambit_actual == 'barcelona':
            continue

        if pd.isna(codi_val) or pd.isna(nom_val):
            continue
        try:
            codi_int = int(float(codi_val))
        except (ValueError, TypeError):
            continue

        nom_str = str(nom_val).strip()

        for tri_label, col_idx in tri_cols.items():
            val = row.iloc[col_idx]
            if pd.isna(val):
                val_net = np.nan
            else:
                try:
                    val_net = float(val)
                except (ValueError, TypeError):
                    val_net = np.nan

            records.append({
                'any': int(any_str),
                'trimestre': tri_label,
                'ambit': ambit_actual,
                'codi': codi_int,
                'nom': nom_str,
                metric_col_name: val_net
            })

    return pd.DataFrame(records)


# ─────────────────────────────────────────────
# 3. LLEGIR I COMBINAR ELS QUATRE XLSX DE GENCAT
# ─────────────────────────────────────────────
def llegeix_xlsx_gencat(path, metric_col_name):
    """
    Llegeix tots els fulls d'un XLSX de Gencat i retorna un DataFrame tidy
    filtrat per a barris (ambit == 'barri').
    """
    print(f"  Llegint {os.path.basename(path)} ...")
    xl = pd.ExcelFile(path)
    dfs = []
    for sheet in xl.sheet_names:
        df = parseja_full_gencat(path, sheet, metric_col_name)
        if not df.empty:
            dfs.append(df)
    if not dfs:
        return pd.DataFrame()
    tot = pd.concat(dfs, ignore_index=True)
    barris = tot[tot['ambit'] == 'barri'].copy()
    barris = barris.drop(columns=['ambit'])
    barris = barris.rename(columns={'codi': 'codi_barri', 'nom': 'nom_barri_raw'})
    barris['nom_barri_norm'] = barris['nom_barri_raw'].apply(normalitza_nom)
    return barris


print("=" * 60)
print("TRACTAMENT DE DADES - PRESSIÓ RESIDENCIAL BCN")
print("=" * 60)

print("\n[1/7] Llegint dades de lloguer (Habitatge Gencat)...")

df_contractes = llegeix_xlsx_gencat(CONTRACTES_PATH, 'contractes_lloguer')
df_lloguer    = llegeix_xlsx_gencat(LLOGUER_PATH,    'lloguer_mitja_eur_mes')
df_lloguer_m2 = llegeix_xlsx_gencat(LLOGUER_M2_PATH, 'lloguer_mitja_eur_m2')
df_sup        = llegeix_xlsx_gencat(SUP_PATH,         'superficie_mitjana_m2')

print(f"  Contractes: {len(df_contractes)} registres de barri-trimestre")
print(f"  Lloguer €/mes: {len(df_lloguer)} registres")
print(f"  Lloguer €/m²: {len(df_lloguer_m2)} registres")
print(f"  Superfície: {len(df_sup)} registres")


# ─────────────────────────────────────────────
# 4. UNIÓ DELS QUATRE DATASETS DE LLOGUER
# ─────────────────────────────────────────────
print("\n[2/7] Unint dades de lloguer per barri-any-trimestre...")

merge_keys = ['codi_barri', 'any', 'trimestre']

df_lloguer_total = (
    df_contractes[merge_keys + ['nom_barri_norm', 'nom_barri_raw', 'contractes_lloguer']]
    .merge(df_lloguer[merge_keys + ['lloguer_mitja_eur_mes']],   on=merge_keys, how='outer')
    .merge(df_lloguer_m2[merge_keys + ['lloguer_mitja_eur_m2']], on=merge_keys, how='outer')
    .merge(df_sup[merge_keys + ['superficie_mitjana_m2']],       on=merge_keys, how='outer')
)

# Recuperem nom_barri_norm per registres que provenen d'outer join sense nom
for src_df in [df_lloguer, df_lloguer_m2, df_sup]:
    mask = df_lloguer_total['nom_barri_norm'].isna()
    if mask.any():
        lookup = src_df.set_index(merge_keys)['nom_barri_norm']
        idxs = df_lloguer_total[mask].set_index(merge_keys).index
        vals = lookup.reindex(idxs).values
        df_lloguer_total.loc[mask, 'nom_barri_norm'] = vals

# Convertim 0 a NaN: valor 0 en lloguer/contractes indica absència de dades publicades
for col in ['contractes_lloguer', 'lloguer_mitja_eur_mes', 'lloguer_mitja_eur_m2', 'superficie_mitjana_m2']:
    if col in df_lloguer_total.columns:
        df_lloguer_total.loc[df_lloguer_total[col] == 0, col] = np.nan

print(f"  Total registres barri-trimestre: {len(df_lloguer_total)}")
print(f"  Anys coberts: {sorted(df_lloguer_total['any'].unique())}")
print(f"  Barris únics: {df_lloguer_total['codi_barri'].nunique()}")


# ─────────────────────────────────────────────
# 5. TAULA MESTRA DE BARRIS (CODIS + NOMS OFICIALS)
# ─────────────────────────────────────────────
print("\n[3/7] Preparant taula mestra de barris...")

df_barris = pd.read_csv(BARRIS_PATH, dtype=str)
df_barris.columns = [c.strip().lower() for c in df_barris.columns]
df_barris['codi_barri_int']     = df_barris['codi_barri'].astype(int)
df_barris['codi_districte_int'] = df_barris['codi_districte'].astype(int)

mestra = df_barris[['codi_barri_int', 'codi_districte_int', 'nom_barri', 'nom_districte']].copy()
mestra.columns = ['codi_barri', 'codi_districte', 'nom_barri', 'nom_districte']

print(f"  Barris al fitxer mestre: {len(mestra)}")

df_lloguer_total = df_lloguer_total.merge(mestra, on='codi_barri', how='left')

no_match = df_lloguer_total['nom_barri'].isna().sum()
if no_match > 0:
    print(f"  AVÍS: {no_match} registres sense correspondència a la taula mestra de barris")


# ─────────────────────────────────────────────
# 6. HABITATGES D'ÚS TURÍSTIC (HUT)
# ─────────────────────────────────────────────
print("\n[4/7] Processant HUT...")

df_hut_raw = pd.read_csv(HUT_PATH)
df_hut_raw.columns = [c.strip() for c in df_hut_raw.columns]

print(f"  Registres HUT bruts: {len(df_hut_raw)}")

# NOTA METODOLÒGICA: el fitxer no conté columna d'estat (actiu/inactiu).
# Es comptabilitzen TOTS els registres del fitxer de comunicació 4T2025.
# El fitxer és la llista oficial de llicències comunicades; es considera que
# tots els registres presents són vigents en aquell moment.

df_hut = (
    df_hut_raw
    .groupby(['CODI_BARRI', 'NOM_BARRI', 'CODI_DISTRICTE', 'NOM_DISTRICTE'])
    .agg(hut_nombre=('N_EXPEDIENT', 'count'),
         hut_places=('NUMERO_PLACES', 'sum'))
    .reset_index()
)
df_hut.columns = [c.lower() for c in df_hut.columns]

print(f"  Barris amb HUT: {len(df_hut)}")
print(f"  Total HUT: {df_hut['hut_nombre'].sum()}")


# ─────────────────────────────────────────────
# 7. POBLACIÓ (PADRÓ 2025)
# ─────────────────────────────────────────────
print("\n[5/7] Processant padró municipal...")

df_padro_raw = pd.read_csv(PADRO_PATH)
df_padro_raw.columns = [c.strip() for c in df_padro_raw.columns]

print(f"  Registres padró bruts: {len(df_padro_raw)} (seccions censals)")

# Sumem per barri (el fitxer conté una fila per secció censal)
df_padro = (
    df_padro_raw
    .groupby(['Codi_Barri', 'Nom_Barri', 'Codi_Districte', 'Nom_Districte'])
    .agg(poblacio=('Valor', 'sum'))
    .reset_index()
)
df_padro.rename(columns={
    'Codi_Barri': 'codi_barri',
    'Nom_Barri': 'nom_barri_padro',
    'Codi_Districte': 'codi_districte_padro',
    'Nom_Districte': 'nom_districte_padro'
}, inplace=True)

print(f"  Barris al padró: {len(df_padro)}")
print(f"  Població total BCN: {df_padro['poblacio'].sum():,.0f}")


# ─────────────────────────────────────────────
# 8. RENDA DISPONIBLE (2022)
# ─────────────────────────────────────────────
print("\n[6/7] Processant renda disponible...")

df_renda_raw = pd.read_csv(RENDA_PATH)
df_renda_raw.columns = [c.strip() for c in df_renda_raw.columns]

print(f"  Registres renda bruts: {len(df_renda_raw)} (seccions censals)")

# NOTA METODOLÒGICA: La renda és per secció censal.
# Usem MITJANA SIMPLE per barri (no ponderada per població) perquè el fitxer de
# padró no conté la distribució per secció censal en el format necessari.
# Una mitjana ponderada seria més precisa però requeriria aparellar secció per secció.
# Limitació: la renda és de 2022; s'utilitza com a variable estructural de context
# socioeconòmic, no com a dada temporal.

df_renda = (
    df_renda_raw
    .groupby(['Codi_Barri', 'Nom_Barri', 'Codi_Districte', 'Nom_Districte'])
    .agg(renda_disponible_pc=('Import_Euros', 'mean'))
    .reset_index()
)
df_renda.rename(columns={
    'Codi_Barri': 'codi_barri',
    'Nom_Barri': 'nom_barri_renda',
    'Codi_Districte': 'codi_districte_renda',
    'Nom_Districte': 'nom_districte_renda'
}, inplace=True)
df_renda['renda_disponible_pc'] = df_renda['renda_disponible_pc'].round(2)

print(f"  Barris amb renda: {len(df_renda)}")
print(f"  Renda mitjana BCN: {df_renda['renda_disponible_pc'].mean():,.0f} €/any/persona")


# ─────────────────────────────────────────────
# 9. CONSTRUCCIÓ DE LA TAULA FINAL BARRI-TRIMESTRE
# ─────────────────────────────────────────────
print("\n[7/7] Construint taula final barri-trimestre...")

df_base = df_lloguer_total[[
    'any', 'trimestre', 'codi_barri', 'codi_districte', 'nom_barri', 'nom_districte',
    'contractes_lloguer', 'lloguer_mitja_eur_mes', 'lloguer_mitja_eur_m2', 'superficie_mitjana_m2'
]].copy()

# Afegim HUT (estàtic: fotografia 4T2025)
df_base = df_base.merge(
    df_hut[['codi_barri', 'hut_nombre', 'hut_places']],
    on='codi_barri', how='left'
)

# Afegim Població (estàtica: padró 2025)
df_base = df_base.merge(
    df_padro[['codi_barri', 'poblacio']],
    on='codi_barri', how='left'
)

# Afegim Renda (estàtica: 2022)
df_base = df_base.merge(
    df_renda[['codi_barri', 'renda_disponible_pc']],
    on='codi_barri', how='left'
)

# ── Variable: hut_per_1000_hab
df_base['hut_per_1000_hab'] = (df_base['hut_nombre'] / df_base['poblacio'] * 1000).round(4)

# ── Variable: ratio_lloguer_renda (lloguer anual / renda per persona)
# Interpreació: fracció de la renda anual que representa el lloguer anual
df_base['ratio_lloguer_renda'] = (
    (df_base['lloguer_mitja_eur_mes'] * 12) / df_base['renda_disponible_pc']
).round(4)

# ── Variació interanual (barri × trimestre)
# Estratègia: per a cada registre (barri, any, trimestre) busquem el valor
# del mateix barri i trimestre a l'any anterior i calculem la variació (%).
ordre_tri = {'I': 1, 'II': 2, 'III': 3, 'IV': 4}
df_base['tri_num'] = df_base['trimestre'].map(ordre_tri)
df_base = df_base.sort_values(['codi_barri', 'any', 'tri_num'])

df_prev = df_base[['codi_barri', 'any', 'trimestre',
                   'lloguer_mitja_eur_mes', 'lloguer_mitja_eur_m2']].copy()
df_prev['any'] = df_prev['any'] + 1  # desplaçament: "any+1" → es creuarà amb l'any actual
df_prev = df_prev.rename(columns={
    'lloguer_mitja_eur_mes': 'lloguer_prev_mes',
    'lloguer_mitja_eur_m2':  'lloguer_prev_m2'
})

df_base = df_base.merge(df_prev, on=['codi_barri', 'any', 'trimestre'], how='left')

df_base['variacio_interanual_lloguer'] = (
    (df_base['lloguer_mitja_eur_mes'] - df_base['lloguer_prev_mes'])
    / df_base['lloguer_prev_mes'] * 100
).round(4)
df_base['variacio_interanual_m2'] = (
    (df_base['lloguer_mitja_eur_m2'] - df_base['lloguer_prev_m2'])
    / df_base['lloguer_prev_m2'] * 100
).round(4)

df_base = df_base.drop(columns=['lloguer_prev_mes', 'lloguer_prev_m2'])

# ── Índex sintètic de pressió residencial
# Normalització min-max (0-1) sobre el conjunt complet de barris i períodes.
# S'utilitza la mitjana de les variables normalitzades disponibles (skipna=True)
# per gestionar valors mancants sense descartar el registre.
# Variables: lloguer €/m², HUT/1000hab, ratio lloguer/renda, variació interanual lloguer.

vars_index = [
    'lloguer_mitja_eur_m2',
    'hut_per_1000_hab',
    'ratio_lloguer_renda',
    'variacio_interanual_lloguer'
]

for v in vars_index:
    vmin = df_base[v].min()
    vmax = df_base[v].max()
    rng  = vmax - vmin
    col_norm = f'{v}_norm'
    if rng > 0:
        df_base[col_norm] = (df_base[v] - vmin) / rng
    else:
        df_base[col_norm] = 0.0

norm_cols = [f'{v}_norm' for v in vars_index]
df_base['index_pressio_residencial'] = df_base[norm_cols].mean(axis=1, skipna=True).round(4)
df_base = df_base.drop(columns=norm_cols)

# ── Categoria de pressió (quartils sobre tota la sèrie temporal)
quartils = df_base['index_pressio_residencial'].quantile([0.25, 0.5, 0.75]).values
q1, q2, q3 = quartils

def classifica_pressio(val):
    if pd.isna(val):
        return np.nan
    if val <= q1:
        return 'baixa'
    elif val <= q2:
        return 'mitjana'
    elif val <= q3:
        return 'alta'
    else:
        return 'molt alta'

df_base['pressio_categoria'] = df_base['index_pressio_residencial'].apply(classifica_pressio)

# ── Ordre final de columnes
cols_finals = [
    'any', 'trimestre',
    'codi_districte', 'nom_districte',
    'codi_barri', 'nom_barri',
    'contractes_lloguer',
    'lloguer_mitja_eur_mes', 'lloguer_mitja_eur_m2', 'superficie_mitjana_m2',
    'hut_nombre', 'hut_places', 'poblacio',
    'renda_disponible_pc',
    'hut_per_1000_hab', 'ratio_lloguer_renda',
    'variacio_interanual_lloguer', 'variacio_interanual_m2',
    'index_pressio_residencial', 'pressio_categoria'
]
cols_finals = [c for c in cols_finals if c in df_base.columns]
df_final = df_base[cols_finals].copy()

print(f"  Registres taula final barri-trimestre: {len(df_final)}")
print(f"  Columnes: {list(df_final.columns)}")


# ─────────────────────────────────────────────
# 10. TAULA FINAL ÚLTIM PERÍODE
# ─────────────────────────────────────────────
df_final['tri_num'] = df_final['trimestre'].map(ordre_tri)
ultim_any     = df_final.dropna(subset=['lloguer_mitja_eur_mes'])['any'].max()
ultim_tri_num = df_final[df_final['any'] == ultim_any].dropna(
    subset=['lloguer_mitja_eur_mes'])['tri_num'].max()
ultim_tri_label = {v: k for k, v in ordre_tri.items()}[ultim_tri_num]

df_ultim = df_final[
    (df_final['any'] == ultim_any) & (df_final['tri_num'] == ultim_tri_num)
].drop(columns=['tri_num']).copy()

df_final = df_final.drop(columns=['tri_num'])

print(f"\n  Últim període disponible: {ultim_any} T{ultim_tri_label}")
print(f"  Registres taula últim període: {len(df_ultim)}")


# ─────────────────────────────────────────────
# 11. GUARDAR SORTIDES
# ─────────────────────────────────────────────
out_tri   = os.path.join(BASE, "taula_final_barris_trimestre.csv")
out_ultim = os.path.join(BASE, "taula_final_barris_ultim_periode.csv")

df_final.to_csv(out_tri,   index=False, encoding='utf-8-sig')
df_ultim.to_csv(out_ultim, index=False, encoding='utf-8-sig')

print(f"\n  Guardat: {out_tri}")
print(f"  Guardat: {out_ultim}")


# ─────────────────────────────────────────────
# 12. RESUM DE QUALITAT DE DADES
# ─────────────────────────────────────────────
def resum_df(nom, df, n_barris_col='codi_barri', observacio=''):
    total = len(df)
    n_cols = len(df.columns)
    pct_nan = df.isnull().mean().mean() * 100 if total > 0 else 0
    n_barris = df[n_barris_col].nunique() if n_barris_col in df.columns else 0
    return {
        'nom_taula': nom,
        'nombre_files': total,
        'nombre_columnes': n_cols,
        'percentatge_valors_mancants': round(pct_nan, 2),
        'nombre_barris_detectats': n_barris,
        'observacions': observacio
    }

resum = [
    resum_df('trimestral_bcn_contractes', df_contractes,
             observacio='Nombre contractes lloguer per barri i trimestre. '
                        'Anys 2014-2025 amb desagregació de barri. '
                        'Valors 0 convertits a NaN (absència de publicació Gencat).'),
    resum_df('trimestral_bcn_lloguer-2', df_lloguer,
             observacio='Lloguer mitjà EUR/mes per barri i trimestre. '
                        'Valors 0 convertits a NaN.'),
    resum_df('trimestral_bcn_lloguer_m2', df_lloguer_m2,
             observacio='Lloguer mitjà EUR/m2/mes per barri i trimestre. '
                        'Valors 0 convertits a NaN.'),
    resum_df('trimestral_bcn_sup', df_sup,
             observacio='Superfície mitjana m2 per barri i trimestre.'),
    resum_df('hut_agregat_barri', df_hut,
             observacio='HUT 4T2025. Sense columna d\'estat actiu: tots els registres comptabilitzats. '
                        'Inclou hut_places (places autoritzades). '
                        'Fotografia estàtica (un valor per barri, no temporal).'),
    resum_df('padro_agregat_barri', df_padro,
             observacio='Padró 2025-01-01. Sumat per barri des de seccions censals. '
                        'Variable estàtica de context.'),
    resum_df('renda_2022_barri', df_renda,
             observacio='Renda 2022 per secció censal, agregada per barri amb MITJANA SIMPLE. '
                        'Any diferent del lloguer: s\'usa com a variable de context. '
                        'IPC no aplicat (fitxer disponible però no integrat per complexitat).'),
    resum_df('taula_final_barris_trimestre', df_final,
             observacio=f'Taula final tidy. '
                        f'Anys {df_final["any"].min()}-{df_final["any"].max()}, '
                        f'{df_final["codi_barri"].nunique()} barris. '
                        f'Index pressió: mitjana de 4 variables normalitzades min-max (skipna). '
                        f'Categoria pressió: per quartils sobre tota la sèrie.'),
    resum_df('taula_final_barris_ultim_periode', df_ultim,
             observacio=f'Últim trimestre disponible: {ultim_any} T{ultim_tri_label}. '
                        f'Una fila per barri. Per a mapa, scatter i rànquing.'),
]

df_resum = pd.DataFrame(resum)
out_resum = os.path.join(BASE, "resum_qualitat_dades.csv")
df_resum.to_csv(out_resum, index=False, encoding='utf-8-sig')
print(f"  Guardat: {out_resum}")


# ─────────────────────────────────────────────
# 13. RESUM FINAL A CONSOLA
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("RESUM FINAL")
print("=" * 60)
print(df_resum[['nom_taula','nombre_files','nombre_barris_detectats',
                'percentatge_valors_mancants']].to_string(index=False))
print(f"\nFitxers generats:")
print(f"  - taula_final_barris_trimestre.csv     ({len(df_final)} files)")
print(f"  - taula_final_barris_ultim_periode.csv ({len(df_ultim)} files)")
print(f"  - resum_qualitat_dades.csv")
print("=" * 60)


# ─────────────────────────────────────────────
# 8. Preparació de la taula d'evolució anual del lloguer per districtes
# ─────────────────────────────────────────────
# Objectiu: generar un CSV en format ample per a Flourish (Line chart).
# Estructura: una fila per any, una columna per districte.
# Valor: lloguer mitjà mensual en euros per districte i any.
#
# Metodologia:
#   - Agrupem per (any, nom_districte) els registres barri-trimestre.
#   - Si existeix 'contractes_lloguer', calculem MITJANA PONDERADA:
#       lloguer_anual = sum(lloguer * contractes) / sum(contractes)
#   - Si no, MITJANA SIMPLE dels trimestres disponibles.
#   - Després fem pivot: files = any, columnes = nom_districte.
# ─────────────────────────────────────────────

print("\n[8] Preparant taula d'evolució anual del lloguer per districtes...")

# Partim de df_final (ja en memòria: taula_final_barris_trimestre)
# Eliminem files sense lloguer ni districte
df_evol = df_final.dropna(subset=['lloguer_mitja_eur_mes', 'nom_districte']).copy()

# ── Agregació trimestral → anual per districte
if 'contractes_lloguer' in df_evol.columns:
    # Descartem files on falti tant el lloguer com els contractes per a la ponderació
    df_evol_valid = df_evol.dropna(subset=['lloguer_mitja_eur_mes', 'contractes_lloguer']).copy()

    # Producte per al numerador de la ponderació
    df_evol_valid['lloguer_x_contractes'] = (
        df_evol_valid['lloguer_mitja_eur_mes'] * df_evol_valid['contractes_lloguer']
    )

    agg_distric = (
        df_evol_valid
        .groupby(['any', 'nom_districte'], as_index=False)
        .agg(
            suma_lloguer_pond=('lloguer_x_contractes', 'sum'),
            suma_contractes=('contractes_lloguer', 'sum'),
            n_trimestres=('trimestre', 'nunique')
        )
    )
    agg_distric['lloguer_mitja_anual'] = (
        agg_distric['suma_lloguer_pond'] / agg_distric['suma_contractes']
    ).round(2)

    metode_agregacio = 'mitjana ponderada per contractes_lloguer'
    print(f"  Mètode: {metode_agregacio}")

    # Advertim si algun districte-any té menys de 4 trimestres
    incomplets = agg_distric[agg_distric['n_trimestres'] < 4]
    if not incomplets.empty:
        print(f"  AVÍS: {len(incomplets)} combinació/ons any-districte amb menys de 4 trimestres:")
        print(incomplets[['any', 'nom_districte', 'n_trimestres']].to_string(index=False))

else:
    # Fallback: mitjana simple dels trimestres disponibles
    agg_distric = (
        df_evol
        .groupby(['any', 'nom_districte'], as_index=False)
        .agg(
            lloguer_mitja_anual=('lloguer_mitja_eur_mes', 'mean'),
            n_trimestres=('trimestre', 'nunique')
        )
    )
    agg_distric['lloguer_mitja_anual'] = agg_distric['lloguer_mitja_anual'].round(2)
    metode_agregacio = 'mitjana simple (contractes_lloguer no disponible)'
    print(f"  Mètode: {metode_agregacio}")

# ── Ordre oficial dels deu districtes de Barcelona
ORDRE_DISTRICTES = [
    'Ciutat Vella',
    'Eixample',
    'Sants-Montjuïc',
    'Les Corts',
    'Sarrià-Sant Gervasi',
    'Gràcia',
    'Horta-Guinardó',
    'Nou Barris',
    'Sant Andreu',
    'Sant Martí',
]

# ── Pivot: files = any, columnes = nom_districte
evolucio_wide = (
    agg_distric
    .pivot(index='any', columns='nom_districte', values='lloguer_mitja_anual')
    .reset_index()
)
evolucio_wide.columns.name = None  # eliminem el nom de l'índex de columnes

# Reordenem les columnes en l'ordre oficial i descartam les no reconegudes
cols_disponibles = [c for c in ORDRE_DISTRICTES if c in evolucio_wide.columns]
cols_no_trobades = [c for c in ORDRE_DISTRICTES if c not in evolucio_wide.columns]
if cols_no_trobades:
    print(f"  AVÍS: districtes no trobats a la taula: {cols_no_trobades}")

evolucio_wide = evolucio_wide[['any'] + cols_disponibles]
evolucio_wide = evolucio_wide.sort_values('any').reset_index(drop=True)

print(f"  Anys resultants: {evolucio_wide['any'].min()}–{evolucio_wide['any'].max()} "
      f"({len(evolucio_wide)} files)")
print(f"  Districtes: {cols_disponibles}")

# ── Guardar
out_evol = os.path.join(BASE, "evolucio_lloguer_districtes_any.csv")
evolucio_wide.to_csv(out_evol, index=False, encoding='utf-8-sig')
print(f"  Guardat: {out_evol}")

# ── Prèvia del fitxer generat
print("\n  Prèvia (primeres i últimes 3 files):")
preview = pd.concat([evolucio_wide.head(3), evolucio_wide.tail(3)])
print(preview.to_string(index=False))
print()
print("  CONFIGURACIÓ FLOURISH (Line chart):")
print("    Labels/time : any")
print(f"    Values      : {', '.join(cols_disponibles)}")
print("    Títol       : Com ha evolucionat el preu del lloguer?")
print("    Subtítol    : Evolució anual del lloguer mitjà mensual per districtes de Barcelona.")


# ─────────────────────────────────────────────
# 9. Preparació de variables normalitzades per al radar chart
# ─────────────────────────────────────────────
# Objectiu: generar variables normalitzades (0-1) de les quatre dimensions
# que composen l'index de pressio residencial, per representar-les
# en un radar chart de Flourish.
#
# El radar chart descompon l'index i permet comparar barris per dimensio:
#   1. Preu del lloguer per m2         (lloguer_m2_norm)
#   2. Densitat HUT per 1.000 habitants (hut_per_1000_norm)
#   3. Ratio lloguer / renda            (ratio_lloguer_renda_norm)
#   4. Variacio interanual del lloguer  (variacio_lloguer_norm)
#
# Metode: normalitzacio min-max sobre el conjunt de barris de l'ultim periode.
# NaN: es conserven. Rang zero: s'assigna 0 (evita divisio per zero).
# ─────────────────────────────────────────────

print("\n[9] Preparant variables normalitzades per al radar chart...")

# Partim de df_ultim (ja en memoria: taula_final_barris_ultim_periode)
df_radar = df_ultim.copy()

# ── Funcio de normalitzacio min-max
def minmax_norm(series):
    """
    Normalitza una serie entre 0 i 1 (min-max).
    - Conserva NaN tal com estan.
    - Si min == max (rang zero), retorna 0 per a tots els valors no nuls.
    """
    min_val = series.min(skipna=True)
    max_val = series.max(skipna=True)
    if pd.isna(min_val) or pd.isna(max_val):
        return series.copy()
    if max_val == min_val:
        return series.where(series.isna(), other=0.0)
    return (series - min_val) / (max_val - min_val)

# ── Mapa: variable original -> nom de la variable normalitzada
variables_norm_radar = {
    'lloguer_mitja_eur_m2':        'lloguer_m2_norm',
    'hut_per_1000_hab':            'hut_per_1000_norm',
    'ratio_lloguer_renda':         'ratio_lloguer_renda_norm',
    'variacio_interanual_lloguer': 'variacio_lloguer_norm',
}

vars_normalitzades_radar = []

for original, nova in variables_norm_radar.items():
    if original in df_radar.columns:
        df_radar[nova] = minmax_norm(df_radar[original]).round(4)
        vars_normalitzades_radar.append(nova)
        n_nan = df_radar[nova].isna().sum()
        print(f"  {nova}: rang [{df_radar[original].min():.4f}, "
              f"{df_radar[original].max():.4f}], NaN={n_nan}")
    else:
        print(f"  AVIS: variable '{original}' no trobada -> '{nova}' no generada")

# ── Columnes de la taula de radar (base + variables normalitzades disponibles)
cols_radar_base = [
    'codi_barri', 'nom_barri', 'nom_districte',
    'pressio_categoria', 'index_pressio_residencial',
]
cols_radar = [c for c in cols_radar_base + vars_normalitzades_radar if c in df_radar.columns]
df_radar_out = df_radar[cols_radar].copy()

# ── Guardar taula completa (73 barris)
out_radar = os.path.join(BASE, "radar_index_pressio_barris.csv")
df_radar_out.to_csv(out_radar, index=False, encoding='utf-8-sig')
print(f"\n  Guardat: {out_radar} ({len(df_radar_out)} barris, {len(cols_radar)} columnes)")

# ─────────────────────────────────────────────
# Seleccio de barris representatius per al radar chart
# ─────────────────────────────────────────────
# El radar chart amb 73 barris seria illegible -> seleccionem 8 barris
# distribuïts per tot l'espectre de l'index (diversitat entre categories).
#
# Estrategia (prioritat A):
#   2 barris "molt alta" (index maxim) + 2 "alta" + 2 "mitjana" + 2 "baixa"
# Fallback (prioritat B, si < 6 barris unics per A):
#   8 barris per quartils de l'index (head(2) per quartil).
# ─────────────────────────────────────────────

df_sel = df_radar_out.dropna(subset=['index_pressio_residencial']).copy()
df_sel = df_sel.sort_values('index_pressio_residencial', ascending=False)

CATS_ORDRE_RADAR = ['molt alta', 'alta', 'mitjana', 'baixa']
N_PER_CAT_RADAR  = 2

seleccio_radar = []
for cat in CATS_ORDRE_RADAR:
    df_cat = df_sel[df_sel['pressio_categoria'] == cat]
    seleccio_radar.append(df_cat.head(N_PER_CAT_RADAR))

df_seleccionats = pd.concat(seleccio_radar).drop_duplicates(subset=['codi_barri'])

# Fallback: quantils si no hem obtingut prou diversitat
if len(df_seleccionats) < 6:
    print("  AVIS: seleccio per categories insuficient -> usant quantils")
    df_sel['_q'] = pd.qcut(
        df_sel['index_pressio_residencial'], q=4,
        labels=['q1', 'q2', 'q3', 'q4'], duplicates='drop'
    )
    df_seleccionats = (
        df_sel.sort_values('index_pressio_residencial', ascending=False)
        .groupby('_q', observed=True).head(2)
        .drop_duplicates(subset=['codi_barri'])
        .drop(columns=['_q'])
    )

df_seleccionats = df_seleccionats.sort_values(
    'index_pressio_residencial', ascending=False
).reset_index(drop=True)

# Guardar taula seleccionada
out_radar_sel = os.path.join(BASE, "radar_index_pressio_barris_seleccionats.csv")
df_seleccionats.to_csv(out_radar_sel, index=False, encoding='utf-8-sig')
print(f"  Guardat: {out_radar_sel} ({len(df_seleccionats)} barris seleccionats)")

# ── Mostra dels barris seleccionats
print("\n  Barris seleccionats per al radar chart:")
cols_show_radar = ['nom_barri', 'nom_districte', 'pressio_categoria',
                   'index_pressio_residencial'] + vars_normalitzades_radar
cols_show_radar = [c for c in cols_show_radar if c in df_seleccionats.columns]
print(df_seleccionats[cols_show_radar].to_string(index=False))

print()
print("  CONFIGURACIO FLOURISH (Radar / Spider chart):")
print("    Fitxer recomanat : radar_index_pressio_barris_seleccionats.csv")
print("    Etiquetes (noms) : nom_barri")
print(f"    Variables (eixos): {', '.join(vars_normalitzades_radar)}")
print("    Color per        : pressio_categoria")


# ─────────────────────────────────────────────
# 10. Conversió de la geometria de barris a WGS84 per a Flourish
# ─────────────────────────────────────────────
# Objectiu: convertir el fitxer de polígons de barris de Barcelona
# (ETRS89/UTM zona 31N, EPSG:25831) a WGS84 (EPSG:4326, lon/lat decimal)
# per poder-lo carregar directament a Flourish o a qualsevol eina web.
#
# El fitxer original conté 1.501 features de diferents unitats administratives
# (terme municipal, districtes, barris, seccions censals, etc.).
# Filtrem TIPUS_UA == 'BARRI' per obtenir els 73 barris de Barcelona.
#
# Camps conservats:
#   codi_barri      (enter 1-73, de la propietat BARRI)
#   nom_barri       (de la propietat NOM)
#   codi_districte  (enter 1-10, de la propietat DISTRICTE)
#
# Nota: el GeoJSON de sortida NO inclou el camp CRS explícit,
# que és el format que espera Flourish (WGS84 per defecte sense CRS).
#
# Dependències: pyproj (pip install pyproj)
# ─────────────────────────────────────────────

print("\n[10] Convertint geometria de barris de EPSG:25831 a WGS84...")

import json as _json
try:
    from pyproj import Transformer as _Transformer

    _POLIGONS_PATH = os.path.join(
        BASE, "Open Data BCN", "BCN_UNITATS_ADM",
        "0301100100_UNITATS_ADM_POLIGONS.json"
    )
    _OUT_GEO = os.path.join(BASE, "barris_barcelona_wgs84.geojson")

    if not os.path.exists(_POLIGONS_PATH):
        print(f"  AVIS: fitxer de polígons no trobat: {_POLIGONS_PATH}")
    else:
        # Transformador EPSG:25831 -> EPSG:4326 (always_xy: x=lon, y=lat)
        _transformer = _Transformer.from_crs("EPSG:25831", "EPSG:4326", always_xy=True)

        def _transform_coords(coords):
            """Converteix llista de [x,y] (UTM31N) a [lon,lat] (WGS84)."""
            return [list(_transformer.transform(c[0], c[1])) for c in coords]

        def _transform_geometry(geom):
            """Converteix recursivament qualsevol geometria Polygon o MultiPolygon."""
            gt = geom['type']
            if gt == 'Polygon':
                return {'type': 'Polygon',
                        'coordinates': [_transform_coords(ring)
                                        for ring in geom['coordinates']]}
            elif gt == 'MultiPolygon':
                return {'type': 'MultiPolygon',
                        'coordinates': [[_transform_coords(ring) for ring in poly]
                                        for poly in geom['coordinates']]}
            else:
                raise ValueError(f"Tipus de geometria no suportat: {gt}")

        with open(_POLIGONS_PATH, encoding='utf-8') as _f:
            _raw = _json.load(_f)

        _features_out = []
        for _ft in _raw['features']:
            _p = _ft['properties']
            if _p['TIPUS_UA'] != 'BARRI':
                continue
            _new_props = {
                'codi_barri':     int(_p['BARRI']),
                'nom_barri':      _p['NOM'],
                'codi_districte': int(_p['DISTRICTE']),
            }
            _features_out.append({
                'type':       'Feature',
                'properties': _new_props,
                'geometry':   _transform_geometry(_ft['geometry'])
            })

        # GeoJSON de sortida sense CRS explícit (Flourish l'interpreta com WGS84)
        _geojson_out = {'type': 'FeatureCollection', 'features': _features_out}
        with open(_OUT_GEO, 'w', encoding='utf-8') as _f:
            _json.dump(_geojson_out, _f, ensure_ascii=False)

        # Validació del rang de coordenades
        _lons, _lats = [], []
        for _ft in _features_out:
            _g = _ft['geometry']
            _rings = _g['coordinates'] if _g['type'] == 'Polygon' \
                     else [r for poly in _g['coordinates'] for r in poly]
            for _r in _rings:
                for _c in _r:
                    _lons.append(_c[0]); _lats.append(_c[1])

        print(f"  Barris convertits: {len(_features_out)} (esperat: 73)")
        print(f"  Longitud: [{min(_lons):.4f}, {max(_lons):.4f}]  (esperat: 2.0-2.3)")
        print(f"  Latitud:  [{min(_lats):.4f}, {max(_lats):.4f}]  (esperat: 41.3-41.5)")
        print(f"  Guardat:  {_OUT_GEO}")
        print()
        print("  CONFIGURACIO FLOURISH (Projection map / Choropleth):")
        print("    Fitxer geometria : barris_barcelona_wgs84.geojson")
        print("    Camp de clau     : codi_barri  (valors 1-73)")
        print("    Camp de nom      : nom_barri")
        print("    Clau a les dades : codi_barri (a taula_final_barris_ultim_periode.csv)")

except ImportError:
    print("  AVIS: pyproj no disponible. Instal·la-la amb: pip install pyproj")
    print("        La conversio de geometria s'ha saltat.")
