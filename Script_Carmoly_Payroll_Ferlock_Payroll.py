
import pandas as pd
import numpy as np
import calendar
import datetime as dt
import re

# =========================================
# 0. Configurações iniciais
# =========================================
df_flex = xl("D4:U29")  # Ajuste conforme o arquivo

plataforma_pago = 'carmoly'
#plataforma_pago = 'ferlock payroll'
cod_moneda_default = '1'

colunas_finais = [
    'External ID', 'Currency', 'Empresa de contrato', 'Plataforma de pago', 'Flex Contable',
    'Importe', 'Subsid. Legal', 'Cuenta', 'Agente', 'Merchant', 'Proyecto', 'Resp. Cargo',
    'Dpto', 'Prod.', 'Tipo Canal', 'Metodo', 'Negocio', 'E. Financiera', 'Marca',
    'Débito', 'Crédito', 'Fecha', 'Periodo', 'Clasificacion Asiento', 'Descripcion',
    'Moneda', 'Nombre del asiento'
]

colunas_valores_base = [
    'TOTAL Líquido Salario Vac.',
    'Aportes Personales BPS e IRPF e IRNR',
    'Aportes Patronales BPS',
    'ANDA',
    'TOTAL PAGO USD'
]

# 1. Plataformas

mapa_plataformas = {
    'carmoly': ('Damiani', 'Carmoly'),
    'ferlock payroll': ('Damiani', 'Ferlock')
}

plataforma_pago_key = plataforma_pago.lower()
if plataforma_pago_key not in mapa_plataformas:
    raise ValueError(f"⚠️ Plataforma '{plataforma_pago}' não encontrada!")

plataforma_pago, empresa_contrato = mapa_plataformas[plataforma_pago_key]

# 2. Fecha, Periodo e Descrição
hoy = dt.datetime.today()
mes_abrev = hoy.strftime('%b')
anio = hoy.year
periodo = f"{mes_abrev} {anio}"
ultimo_dia = calendar.monthrange(anio, hoy.month)[1]
fecha = f"{ultimo_dia:02d}/{hoy.month:02d}/{anio}"
nombre_del_asiento = f"{hoy.day:02d}_{hoy.month:02d} Pagos {periodo} {plataforma_pago} {empresa_contrato}"

# 3. Limpeza inicial e colunas duplicadas
def normalizar_nome_coluna(col):
    col = str(col).strip()
    if col in ['nan', 'None', '']:
        return None
    col = re.sub(r'\s+', ' ', col) 
    col = re.sub(r'[A-Za-z]{3}\.\d{4}$', '', col)  # remove "Abr.2025"
    return col.strip()

df_flex.columns = [normalizar_nome_coluna(c) for c in df_flex.iloc[0]]
df_flex = df_flex.drop(index=0).reset_index(drop=True)
df_flex = df_flex.loc[:, df_flex.columns.notna()]
df_flex['FLEX'] = df_flex['FLEX'].astype(str).str.strip()
df_flex = df_flex[df_flex['FLEX'] != '']

# 4. Normalização numérica
for col in df_flex.columns:
    if col not in ['FLEX', 'Fecha', 'Detalle', 'Moneda']:
        df_flex[col] = (
            df_flex[col].astype(str).str.upper()
            .str.replace('USD', '', regex=True)
            .str.replace(',', '.', regex=True)
            .replace({'NONE': '0', '': '0'})     # substitui valores inválidos
            .astype(float)

        )
        df_flex[col] = pd.to_numeric(df_flex[col], errors='coerce')

# 5. Normaliza moeda
mapa_monedas = {
    'UYU': 1, 'GBP': 2, 'CAD': 3, 'EUR': 4, 'USD': 5, 'ARS': 6, 'BOB': 7, 'BRL': 8,
    'XOF': 9, 'CLP': 10, 'CNY': 11, 'COP': 12, 'GHS': 13, 'CRC': 14, 'IDR': 15
}
mapa_codigos = {str(v): k for k, v in mapa_monedas.items()} 

def normalizar_moeda(x):
    x = str(x).strip().upper()
    if x.isdigit():
        codigo = x
        sigla = mapa_codigos.get(codigo, mapa_codigos[str(cod_moneda_default)])
    else:
        sigla = x
        codigo = str(mapa_monedas.get(sigla, cod_moneda_default))
    return pd.Series([sigla, codigo])

if 'Moneda' not in df_flex.columns:
    df_flex['Moneda'] = cod_moneda_default

df_flex[['Currency', 'Moneda']] = df_flex['Moneda'].apply(normalizar_moeda)

# Fecha e Periodo
if 'Fecha' not in df_flex.columns:
    df_flex['Fecha'] = fecha
df_flex['Fecha'] = pd.to_datetime(df_flex['Fecha'], errors='coerce').fillna(pd.Timestamp.today())
df_flex['Fecha'] = df_flex['Fecha'].dt.strftime('%d/%m/%Y')
df_flex['Periodo'] = periodo

# 6. Mapeamento de cuentas crédito
mapa_cuentas_credito = {
    'TOTAL Líquido Salario Vac.': '21211.0000',
    'TOTAL PAGO USD': '21211.0000',
    'Aportes Personales BPS e IRPF e IRNR': '21212.0500',
    'Aportes Patronales BPS': '21212.0500',
    'ANDA': '21212.0502'
}

# 7. Processamento final
df_final_total = pd.DataFrame(columns=colunas_finais)

for col_valor in colunas_valores_base:
    col_match = [c for c in df_flex.columns if col_valor.lower() in c.lower()]

    if not col_match:
        continue

    # --- Calcula Importe ---
    if col_valor == 'ANDA':
        col_bse = [c for c in df_flex.columns if 'total importe bse' in c.lower()]
        df_flex['Importe'] = (
            pd.to_numeric(df_flex[col_match[0]], errors='coerce').fillna(0) +
            pd.to_numeric(df_flex[col_bse[0]], errors='coerce').fillna(0) if col_bse else
            pd.to_numeric(df_flex[col_match[0]], errors='coerce').fillna(0)
        )
    else:
        df_flex['Importe'] = pd.to_numeric(df_flex[col_match[0]], errors='coerce').fillna(0)

    # 🔹 Força USD no TOTAL PAGO USD
    if col_valor == 'TOTAL PAGO USD':
        df_flex['Moneda'] = '8'
        df_flex['Currency'] = 'USD'

    # --- Cria df_melt ---
    df_melt = df_flex[['FLEX','Moneda','Fecha','Periodo','Currency','Importe']].copy()
    df_melt = df_melt[df_melt['Importe'] != 0]

    if df_melt.empty:
        continue

    # 🔹 Quebra FLEX (com regra fixa para nomes)
    mask_letras = df_melt['FLEX'].str.contains('[A-Za-z]', regex=True)
    flex_parts = pd.DataFrame(index=df_melt.index)

    for i in range(13):
        flex_parts[i] = ''  # inicializa vazio

    for idx, val in df_melt['FLEX'].items():
        if mask_letras[idx]:
            flex_parts.loc[idx] = [
                '21','11415.0004','0000','00000','0103','996','1101','1400','00','00','0000','0000','0000'
            ]
        else:
            parts = val.split('_')
            for i, p in enumerate(parts[:13]):
                flex_parts.loc[idx, i] = p

    colunas_detalhadas = [
        'Subsid. Legal','Cuenta','Agente','Merchant','Proyecto','Resp. Cargo',
        'Dpto','Prod.','Tipo Canal','Metodo','Negocio','E. Financiera','Marca'
    ]

    # --- Bloco Débito ---
    df_debito = df_melt.copy()
    df_debito[colunas_detalhadas] = flex_parts.values
    df_debito['Débito'] = df_debito['Importe']
    df_debito['Crédito'] = 0
    df_debito.loc[mask_letras, 'Cuenta'] = '11415.0004'

    # --- Bloco Crédito ---
    df_credito = df_melt.copy()
    df_credito[colunas_detalhadas] = flex_parts.values
    df_credito['Débito'] = 0
    df_credito['Crédito'] = df_credito['Importe']
    if col_valor in mapa_cuentas_credito:
        df_credito['Cuenta'] = mapa_cuentas_credito[col_valor]

    # --- Concatena blocos ---
    df_block = pd.concat([df_debito, df_credito], ignore_index=True)
    df_block['External ID'] = f"Nom_Pagos_{mes_abrev}{anio}"
    df_block['Empresa de contrato'] = empresa_contrato
    df_block['Plataforma de pago'] = plataforma_pago
    df_block['Flex Contable'] = df_block['FLEX']
    df_block['Clasificacion Asiento'] = 'CSV Otras Reclasificaciones'
    df_block['Descripcion'] = (
        f"{nombre_del_asiento} | {plataforma_pago} | {empresa_contrato} | " + df_block['Currency'].astype(str)
    )
    df_block['Nombre del asiento'] = nombre_del_asiento
    df_block = df_block[colunas_finais]

    # Linha total
    total_debito = df_block['Débito'].sum()
    total_credito = df_block['Crédito'].sum()
    linha_total = {col: '' for col in df_block.columns}
    linha_total['Débito'] = total_debito
    linha_total['Crédito'] = total_credito
    linha_total['Fecha'] = 'VERDADEIRO' if total_debito == total_credito else 'FALSO'
    df_block = pd.concat([df_block, pd.DataFrame([linha_total])], ignore_index=True)

    df_final_total = pd.concat([df_final_total, df_block], ignore_index=True)

# 8. Resultado final
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
df_final_total = df_final_total[colunas_finais]