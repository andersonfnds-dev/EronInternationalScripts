import pandas as pd
import numpy as np
import calendar
import datetime as dt

# =========================================
# 0. Configurações iniciais
# =========================================
df_flex = xl("A6:Z100")  # Ajuste conforme o arquivo
plataforma_pago = 'liteup payroll'
cod_moneda_default = '22'

colunas_finais = [
    'External ID', 'Currency', 'Empresa de contrato', 'Plataforma de pago', 'Flex Contable',
    'Importe', 'Subsid. Legal', 'Cuenta', 'Agente', 'Merchant', 'Proyecto', 'Resp. Cargo',
    'Dpto', 'Prod.', 'Tipo Canal', 'Metodo', 'Negocio', 'E. Financiera', 'Marca',
    'Débito', 'Crédito', 'Fecha', 'Periodo', 'Clasificacion Asiento', 'Descripcion',
    'Moneda', 'Nombre del asiento'
]

colunas_valores_base = ['TOTAL INGRESOS', 'Total AFP', 'RTA 5TA', 'Neto A Pagar', 'Essalud']

mapa_plataformas = {
    'op': ('OP Payroll', 'OP'),
}

plataforma_pago_key = plataforma_pago.lower()
if plataforma_pago_key not in mapa_plataformas:
    raise ValueError(f"⚠️ Plataforma '{plataforma_pago}' não encontrada!")

plataforma_pago, empresa_contrato = mapa_plataformas[plataforma_pago_key]

# =========================================
# 2. Fecha, Periodo e Descrição
# =========================================
hoy = dt.datetime.today()
mes_abrev = hoy.strftime('%b')
anio = hoy.year
periodo = f"{mes_abrev} {anio}"
ultimo_dia = calendar.monthrange(anio, hoy.month)[1]
fecha = f"{ultimo_dia:02d}/{hoy.month:02d}/{anio}"

nombre_del_asiento = f"{hoy.day:02d}_{hoy.month:02d} Pagos {periodo} {plataforma_pago} {empresa_contrato}"

# =========================================
# 3. Limpeza inicial
# =========================================
df_flex.columns = [str(c).strip() for c in df_flex.iloc[0]]
df_flex = df_flex.drop(index=0).reset_index(drop=True)

# Ajusta a coluna NOMBRE como FLEX
if 'NOMBRE' not in df_flex.columns:
    raise ValueError("⚠️ Coluna 'NOMBRE' não encontrada na planilha!")

df_flex.rename(columns={'NOMBRE': 'FLEX'}, inplace=True)
df_flex['FLEX'] = df_flex['FLEX'].astype(str).str.strip()
df_flex = df_flex[~df_flex['FLEX'].isin(['', 'None', 'nan', 'NaN'])].copy()

mapa_monedas = {
    'UYU': 1, 'GBP': 2, 'CAD': 3, 'EUR': 4, 'USD': 5, 'ARS': 6, 'BOB': 7, 'BRL': 8,
    'XOF': 9, 'CLP': 10, 'CNY': 11, 'COP': 12, 'GHS': 13, 'CRC': 14, 'IDR': 15,
    'INR': 16, 'JPY': 17, 'KES': 18, 'MXN': 19, 'MYR': 20, 'NGN': 21, 'PEN': 22
}
mapa_codigos = {str(v): k for k, v in mapa_monedas.items()} 

def normalizar_moeda(x):
    x = str(x).strip().upper()
    if x.isdigit():
        sigla = mapa_codigos.get(x, mapa_codigos[str(cod_moneda_default)])
        codigo = x
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

# =========================================
# 4. Processamento geral
# =========================================
df_final_total = pd.DataFrame(columns=colunas_finais)
mapa_cuentas_credito = {
    'total ingresos': '21211.0000',
    'total afp': '21222.0000',
    'rta 5ta': '21223.0000',
    'neto a pagar': '21224.0000',
    'essalud': '21225.0000'
}

def processar_dataframe(df_base):
    df_final = pd.DataFrame(columns=colunas_finais)
    for col_valor in colunas_valores_base:
        col_match = [c for c in df_base.columns if col_valor.lower() in c.lower()]
        if not col_match:
            continue

        df_melt = df_base[['FLEX', 'Moneda', 'Fecha', 'Periodo', 'Currency'] + col_match].copy()
        df_melt['Importe'] = pd.to_numeric(df_melt[col_match[0]], errors='coerce')
        df_melt = df_melt.dropna(subset=['Importe'])
        df_melt = df_melt[df_melt['Importe'] != 0]
        if df_melt.empty:
            continue

        # Quebra FLEX em 13 partes
        flex_parts = df_melt['FLEX'].str.split('_', n=12, expand=True)
        while flex_parts.shape[1] < 13:
            flex_parts[flex_parts.shape[1]] = ''
        flex_parts = flex_parts.iloc[:, :13]

        colunas_detalhadas = [
            'Subsid. Legal','Cuenta','Agente','Merchant','Proyecto','Resp. Cargo',
            'Dpto','Prod.','Tipo Canal','Metodo','Negocio','E. Financiera','Marca'
        ]
        df_melt[colunas_detalhadas] = flex_parts.values

        # Débito
        df_debito = df_melt.copy()
        df_debito['Débito'] = df_debito['Importe']
        df_debito['Crédito'] = 0

        # Crédito
        df_credito = df_melt.copy()
        df_credito['Débito'] = 0
        df_credito['Crédito'] = df_credito['Importe']
        conta_credito = next((c for k, c in mapa_cuentas_credito.items() if k in col_valor.lower()), '99999.0000')
        df_credito['Cuenta'] = conta_credito

        # Junta e formata
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

        # Linha de total
        total_debito = df_block['Débito'].sum()
        total_credito = df_block['Crédito'].sum()
        linha_total = {col: '' for col in df_block.columns}
        linha_total['Débito'] = total_debito
        linha_total['Crédito'] = total_credito
        linha_total['Fecha'] = 'VERDADEIRO' if total_debito == total_credito else 'FALSO'
        df_block = pd.concat([df_block, pd.DataFrame([linha_total])], ignore_index=True)

        df_final = pd.concat([df_final, df_block], ignore_index=True)

    return df_final

df_final_total = processar_dataframe(df_flex)

# =========================================
# 5. Exibição final
# =========================================
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
df_final_total = df_final_total[colunas_finais]
df_final_total
