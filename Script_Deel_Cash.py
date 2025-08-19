import calendar
import datetime as dt
import pandas as pd

# =========================================
# ⚠ CONFIGURACIÓN MANUAL ⚠
# Ajustar rango según el archivo
# BORRE el símbolo "#" en la línea correspondiente a plataforma_pago
# Ajustar cod_moneda según la moneda utilizada

# -----------------------------------------

df_flex = pd.read_excel('C:\\Users\\Usuario\\Downloads\\Info Excel Solutions\\Info Excel Solutions\\CASH\\Listo - FLEX - Cash (Abril).xlsx')  # Ejemplo: xl("A1:D5")

plataforma_pago = input("Qual a plataforma de pago? (cash, deel, op, damiani, d24, liteup payroll, liteup contractors, luzino payroll, luzino contractors, carmoly, ferlock) ")

cod_moneda = input("Qual o código da moeda? (ex: USD, EUR, UYU) Deixe vazio se a coluna Moneda estiver preenchida: ")
#plataforma_pago = 'd24'
#plataforma_pago = 'luzino contractors'
#plataforma_pago = 'ferlock payroll'
#plataforma_pago = 'ferlock contractors'
#plataforma_pago = 'op'
#plataforma_pago = 'cash'
#plataforma_pago = 'damiani'
#plataforma_pago = 'liteup payroll'
#plataforma_pago = 'liteup contractors'
#plataforma_pago = 'luzino payroll'
#plataforma_pago = 'carmoly'

#cod_moneda = '5'  
# Preencha caso a coluna Moneda esteja totalmente vazia
# =========================================
# 1. Mapeo de plataformas y empresas
# =========================================
mapa_tipos_tabelas = {
    'tipo1' : ('Cash', 'Deel', 'LiteUp Contractors'),
    'tipo2' : ('D24','Luzino Contractors','Ferlock Contractors')
}

mapa_plataformas = {
    'cash': ('Cash', 'Luzino'),
    'op': ('OP Payroll', 'OP'),
    'deel': ('Deel', 'Luzino'),
    'damiani': ('Damiani', 'Directa 24 LLC'),
    'd24': ('Damiani', 'Directa 24 LLC'),
    'liteup payroll': ('Payroll', 'LiteUp'),
    'liteup contractors': ('Contractors', 'LiteUp'),
    'luzino payroll': ('Payroll', 'Luzino'),
    'luzino contractors': ('Damiani', 'Luzino'),
    'carmoly': ('Damiani', 'Carmoly'),
    'ferlock': ('Damiani', 'Ferlock')
}

if plataforma_pago.lower() not in mapa_plataformas:
    raise ValueError(
        f"⚠ Plataforma '{plataforma_pago}' no encontrada.\n"
        f"Opciones válidas: {', '.join(mapa_plataformas.keys())}"
    )

plataforma_pago_key = plataforma_pago.lower()
plataforma_pago, empresa_contrato = mapa_plataformas[plataforma_pago_key]

# =========================================
# 2. Fecha y Período
# =========================================
hoy = dt.datetime.today()
mes_abrev = hoy.strftime('%b')
anio = hoy.year
periodo = f"{mes_abrev} {anio}"
ultimo_dia = calendar.monthrange(anio, hoy.month)[1]
fecha = f"{ultimo_dia:02d}/{hoy.month:02d}/{anio}"

nombre_del_asiento = f"{hoy.day:02d}_{hoy.month:02d} Pagos {periodo} {plataforma_pago} {empresa_contrato}"

# Dicionário ISO -> ID interno (Moneda)
mapa_monedas = {
    'UYU': 1, 'GBP': 2, 'CAD': 3, 'EUR': 4, 'USD': 5, 'ARS': 6, 'BOB': 7, 'BRL': 8,
    'XOF': 9, 'CLP': 10, 'CNY': 11, 'COP': 12, 'GHS': 13, 'CRC': 14, 'IDR': 15,
    'INR': 16, 'JPY': 17, 'KES': 18, 'MXN': 19, 'MYR': 20, 'NGN': 21, 'PEN': 22,
    'PHP': 23, 'PYG': 24, 'THB': 25, 'TZS': 26, 'VES': 27, 'VND': 28, 'XAF': 29
}

# =========================================
# 3. Verificación de Moneda
# =========================================
colunas = list(df_flex.iloc[0])
is_novo_layout = 'Detalle' in colunas and 'Debe' in colunas

# # df_flex.columns = colunas
# # print(df_flex.columns)
# df_flex = df_flex.drop(index=0).reset_index(drop=True)
# df_flex.columns = [str(col).strip() for col in df_flex.columns]

df_flex['Moneda'] = df_flex['Moneda'].astype(str).str.strip()
df_flex['Moneda'].replace({'nan': '', 'NaN': ''}, regex=False, inplace=True)

print(df_flex)

if df_flex['Moneda'].eq('').all():
    # Coluna Moneda está completamente vazia
    if cod_moneda == '':
        raise ValueError("⚠ Nenhuma moeda foi detectada. Preencha 'cod_moneda'.")
    df_flex['Currency'] = cod_moneda.upper()
    df_flex['Moneda'] = mapa_monedas.get(cod_moneda.upper(), cod_moneda.upper())
else:
    # Coluna Moneda preenchida
    df_flex['Moneda'] = df_flex['Moneda'].str.upper()
    invalidos = [m for m in df_flex['Moneda'].unique() if m not in mapa_monedas]
    if cod_moneda == '':
        pass
    df_flex['Currency'] = df_flex['Moneda']           # ISO visível
    df_flex['Moneda'] = df_flex['Moneda'].map(mapa_monedas)  # Código interno

# =========================================
# 4. Limpeza e tratamento do layout
# =========================================

if 'Pais' in df_flex.columns:
    df_flex = df_flex.drop(columns=['Pais'])

col_honorario = next((col for col in df_flex.columns if 'honorario' in col.casefold()), None)
if not col_honorario:
    raise ValueError("No se encontró una columna de honorarios.")
df_flex = df_flex.rename(columns={col_honorario: 'Honorarios'})

df_flex = df_flex.dropna(subset=['FLEX', 'Honorarios'])
df_flex['Honorarios'] = pd.to_numeric(df_flex['Honorarios'], errors='coerce')

if is_novo_layout:
    df_flex['Fecha'] = pd.to_datetime(df_flex['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y')
else:
    df_flex['Fecha'] = fecha

# Consolidar FLEX duplicados
df_flex = df_flex.groupby(['FLEX', 'Moneda', 'Fecha', 'Currency'], as_index=False)['Honorarios'].sum()

# Expandir FLEX
df_flex['Subsid. Legal'] = ''
colunas_detalhadas = [
    'Cuenta', 'Agente', 'Merchant', 'Proyecto', 'Resp. Cargo', 'Dpto',
    'Prod.', 'Tipo Canal', 'Metodo', 'Negocio', 'E. Financiera', 'Marca'
]
for col in colunas_detalhadas:
    df_flex[col] = ''

for i, row in df_flex.iterrows():
    partes = str(row['FLEX']).split('_')
    if len(partes) == 13:
        df_flex.at[i, 'Subsid. Legal'] = partes[0]
        for j, col in enumerate(colunas_detalhadas):
            df_flex.at[i, col] = partes[j + 1]

# =========================================
# 5. Débitos y Créditos
# =========================================
df_debito = df_flex.copy()
df_credito = df_flex.copy()

df_debito['Débito'] = df_debito['Honorarios']
df_debito['Crédito'] = 0
df_credito['Débito'] = 0
df_credito['Crédito'] = df_credito['Honorarios']

if plataforma_pago_key in ['deel', 'cash', 'liteup contractors']:
    df_credito['Cuenta'] = '21211.0000'  # Substituir conta crédito
else:
    df_credito['Cuenta'] = df_credito['FLEX'].str.split('_').str[0]  # Mantém conta original

df_final = pd.concat([df_debito, df_credito], ignore_index=True)

# Campos fixos
df_final['Período'] = periodo
df_final['Clasificacion Asiento'] = 'CSV Otras Reclasificaciones'
df_final['Nombre del asiento'] = nombre_del_asiento
df_final['Descripcion'] = (
    nombre_del_asiento + ' | ' + plataforma_pago + ' | ' + empresa_contrato + ' | ' + df_final['Currency']
)
df_final['Empresa de contrato'] = empresa_contrato
df_final['Plataforma de pago'] = plataforma_pago
df_final['Flex Contable'] = df_final['FLEX']
df_final['Importe'] = df_final['Honorarios']

# =========================================
# 6. Orden y ID externo
# =========================================
colunas_finais = [
    'External ID', 'Currency', 'Empresa de contrato', 'Plataforma de pago', 'Flex Contable',
    'Importe', 'Subsid. Legal', 'Cuenta', 'Agente', 'Merchant', 'Proyecto', 'Resp. Cargo',
    'Dpto', 'Prod.', 'Tipo Canal', 'Metodo', 'Negocio', 'E. Financiera', 'Marca',
    'Débito', 'Crédito', 'Fecha', 'Período', 'Clasificacion Asiento', 'Descripcion',
    'Moneda', 'Nombre del asiento'
]

df_final['External ID'] = f"Nom_Pagos_{mes_abrev}{anio}"

total_debito = df_final['Débito'].sum()
total_credito = df_final['Crédito'].sum()
valido = total_debito == total_credito

linha_total = {col: '' for col in df_final.columns}
linha_total['Débito'] = total_debito
linha_total['Crédito'] = total_credito
linha_total['Fecha'] = 'VERDADEIRO' if valido else 'FALSO'

df_final = pd.concat([df_final, pd.DataFrame([linha_total])], ignore_index=True)

# Ajustes finais
df_final['Período'] = df_final['Período'].astype(str)
df_final['Fecha'] = df_final['Fecha'].apply(
    lambda x: x if x in ['VERDADEIRO', 'FALSO'] else pd.to_datetime(x, errors='coerce', dayfirst=True).strftime('%d/%m/%Y')
)

# Mostrar no Excel
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
df_final = df_final[colunas_finais]
print(df_final)