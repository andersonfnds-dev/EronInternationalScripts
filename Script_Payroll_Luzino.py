import pandas as pd
import numpy as np
import calendar
import datetime as dt

# =========================================
# 0. Configurações iniciais
# =========================================
df_raw = xl("A6:M30", header=None, dtype=str)  # lê tudo como string

plataforma_pago = 'luzino payroll'
cod_moneda_default = '4'

colunas_finais = [
    'External ID', 'Currency', 'Empresa de contrato', 'Plataforma de pago', 'Flex Contable',
    'Importe', 'Subsid. Legal', 'Cuenta', 'Agente', 'Merchant', 'Proyecto', 'Resp. Cargo',
    'Dpto', 'Prod.', 'Tipo Canal', 'Metodo', 'Negocio', 'E. Financiera', 'Marca',
    'Débito', 'Crédito', 'Fecha', 'Periodo', 'Clasificacion Asiento', 'Descripcion',
    'Moneda', 'Nombre del asiento'
]

# 1. Plataformas
mapa_plataformas = {
    'luzino payroll': ('Payroll', 'Luzino'),
}

plataforma_pago_key = plataforma_pago.lower()
if plataforma_pago_key not in mapa_plataformas:
    raise ValueError(f"⚠️ Plataforma '{plataforma_pago}' não encontrada!")

empresa_contrato, plataforma_pago_fmt = mapa_plataformas[plataforma_pago_key]

# =========================================
# 2. Fecha, Periodo e Descrição
# =========================================
hoy = dt.datetime.today()
mes_abrev = hoy.strftime('%b')
anio = hoy.year
periodo = f"{mes_abrev} {anio}"
ultimo_dia = calendar.monthrange(anio, hoy.month)[1]
fecha = f"{ultimo_dia:02d}/{hoy.month:02d}/{anio}"
nombre_del_asiento = f"{hoy.day:02d}_{hoy.month:02d} Pagos {periodo} {plataforma_pago_fmt} {empresa_contrato}"

# =========================================
# 3. Ajusta nomes de colunas
# =========================================
colunas_iniciais = ['CONCEPTO', 'INFO', 'TOTAL']
flex_headers = df_raw.iloc[0, 3:].tolist()  # headers reais dos Flex
df_raw.columns = colunas_iniciais + flex_headers
df_raw = df_raw.iloc[1:].reset_index(drop=True)  # remove a linha de headers duplicada

col_total_idx = 2  # "TOTAL" é a 3ª coluna
flex_cols = df_raw.columns[col_total_idx+1:]  # todas as colunas de Flex

# Localiza linhas-chave
linha_total_ret = df_raw.index[
    df_raw.iloc[:,0].str.contains("TOTAL RETENCION", case=False, na=False)
][0]

linha_total_coste = df_raw.index[
    df_raw.iloc[:,0].str.contains("TOTAL COSTE S.S. EMPRESA", case=False, na=False)
][0]

# =========================================
# 4. Função para quebrar Flex em 13 partes
# =========================================
def quebrar_flex(flex):
    parts = str(flex).split('_')
    parts = parts[:13] + ['']*(13-len(parts))
    return parts

colunas_detalhadas = [
    'Subsid. Legal','Cuenta','Agente','Merchant','Proyecto','Resp. Cargo',
    'Dpto','Prod.','Tipo Canal','Metodo','Negocio','E. Financiera','Marca'
]

# =========================================
# 5. Normalização de moedas
# =========================================
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

# =========================================
# 6. Montagem do DataFrame Final
# =========================================
df_final = pd.DataFrame(columns=colunas_finais)

def get_float_value(df, row_idx, col):
    val = df.loc[row_idx, col]
    if isinstance(val, pd.Series):
        val = val.iloc[0]
    # Converte para string antes de fazer replace
    val_str = str(val).strip()
    # Ignora se for vazio ou NaN
    if val_str in ['', 'nan', 'None']:
        return 0.0
    return float(val_str.replace(',', '.'))

for col in flex_cols:
    flex_name = col.strip()
    valores_padrao = 100.0
    valores_retencion = get_float_value(df_raw, linha_total_ret, col)
    valores_coste = get_float_value(df_raw, linha_total_coste, col)
    
    valor_total = valores_padrao + valores_retencion + valores_coste

     # Quebra o Flex uma vez só
    parts = quebrar_flex(flex_name)
    cuenta_flex = parts[1]  # A conta vem da segunda posição do Flex
    
    registros = [
        (valores_padrao, '21211.0000', 0, valores_padrao),  # Crédito
        (valores_retencion, '21212.0000', 0, valores_retencion),  # Crédito
        (valores_coste, '21212.0000', 0, valores_coste),  # Crédito
        (valor_total, cuenta_flex, valor_total, 0)  # Débito final
    ]
    
    for valor, cuenta, debito, credito in registros:
        parts = quebrar_flex(flex_name)
        row = {
            'External ID': f"Nom_Pagos_{mes_abrev}{anio}",
            'Currency': 'USD',  # será sobrescrito após normalização
            'Empresa de contrato': empresa_contrato,
            'Plataforma de pago': plataforma_pago_fmt,
            'Flex Contable': flex_name,
            'Importe': valor,
            **dict(zip(colunas_detalhadas, parts)),
            'Cuenta': cuenta,
            'Débito': debito,
            'Crédito': credito,
            'Fecha': fecha,
            'Periodo': periodo,
            'Clasificacion Asiento': 'CSV Otras Reclasificaciones',
            'Descripcion': f"{nombre_del_asiento} | {plataforma_pago_fmt} | {empresa_contrato} | USD",
            'Moneda': cod_moneda_default,
            'Nombre del asiento': nombre_del_asiento
        }
        df_final.loc[len(df_final)] = row

# =========================================
# 7. Normaliza Moeda, Fecha e Periodo
# =========================================
df_final[['Currency','Moneda']] = df_final['Moneda'].apply(normalizar_moeda)
df_final['Fecha'] = pd.to_datetime(df_final['Fecha'], errors='coerce').fillna(pd.Timestamp.today())
df_final['Fecha'] = df_final['Fecha'].dt.strftime('%d/%m/%Y')
df_final['Periodo'] = periodo

# =========================================
# 8. Linha total de controle
# =========================================
total_debito = df_final['Débito'].sum()
total_credito = df_final['Crédito'].sum()

linha_total = {col: '' for col in df_final.columns}
linha_total['Débito'] = total_debito
linha_total['Crédito'] = total_credito
linha_total['Fecha'] = 'VERDADEIRO' if total_debito == total_credito else 'FALSO'

df_final = pd.concat([df_final, pd.DataFrame([linha_total])], ignore_index=True)

# =========================================
# 9. Resultado final
# =========================================
df_final = df_final[colunas_finais]
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
df_final