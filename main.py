import pandas as pd
import plotly.express as px
import unicodedata
import seaborn as sns
import matplotlib.pyplot as plt
# ===============================
# 1. ARCHIVO Y HOJA
# ===============================
archivo = "DatosSalida.xlsx"
hoja = "BASE PUERTO SALIDA"

PALETA_MORADO = {
    'Trips': '#6A0DAD',     # morado fuerte
    'Afidos': '#B19CD9'     # morado claro
}
PALETA_MORADO3 = [
    '#6A0DAD',  # morado fuerte
    '#B19CD9' ]   # morado claro

ORDEN_BLANCOS = ["Trips", "Afidos"]

# ===============================

# 2. CARGAR SIN HEADER
# ===============================
raw = pd.read_excel(archivo, sheet_name=hoja, header=None)

# ===============================
# 3. DETECTAR FILA DE ENCABEZADOS
# ===============================
header_row = None
for i in range(len(raw)):
    fila = raw.iloc[i].astype(str).str.upper().tolist()
    if "PRODUCTO" in fila and "AÑO" in fila:
        header_row = i
        break

if header_row is None:
    raise ValueError("❌ No se encontró la fila de encabezados")

print(f"✅ Encabezados encontrados en la fila {header_row}")

# ===============================
# 4. RELEER CON HEADER CORRECTO
# ===============================
df = pd.read_excel(archivo, sheet_name=hoja, header=header_row)

# ===============================
# 5. LIMPIEZA DE COLUMNAS
# ===============================
def limpiar_texto(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    x = unicodedata.normalize('NFKD', x)
    x = x.encode('ascii', 'ignore').decode('utf-8')
    return x.title()

def limpiar_columnas(df):
    cols = []
    for col in df.columns:
        col = unicodedata.normalize('NFKD', str(col))
        col = col.encode('ascii', 'ignore').decode('utf-8')
        col = col.strip().lower().replace(" ", "_")
        cols.append(col)
    df.columns = cols
    return df

df = limpiar_columnas(df)

# ===============================
# 6. LIMPIEZA GENERAL
# ===============================
campos_texto = [
    'predio',
    'poscosecha_proceso',
    'pais',
    'cliente',
    'blanco_biologico'
]

for c in campos_texto:
    df[c] = df[c].apply(limpiar_texto)

# eliminar registros inválidos
invalidos = ['No', 'N/A', 'None', '']
for c in campos_texto:
    df = df[~df[c].isin(invalidos)]

# ===============================
# 7. UNIFICAR CLIENTES (ABCO)
# ===============================
def normalizar_cliente(c):
    if c is None:
        return None
    if "Abco" in c:
        return "Distribuidora Abco S.A"
    return c

df['cliente'] = df['cliente'].apply(normalizar_cliente)

# ===============================
# 8. FECHAS Y NÚMEROS
# ===============================
if 'fecha' in df.columns:
    df['fecha'] = pd.to_datetime(df['fecha'], dayfirst=True, errors='coerce')

cols_num = ['cuenta', 'cuenta_producto', 'total_piezas', 'total_tallos_rechazados']
for c in cols_num:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

# ===============================
# 9. NORMALIZAR BLANCO BIOLÓGICO
# ===============================
def normalizar_blanco(v):
    if v is None:
        return "OTROS"
    v = v.upper()

    if "ACAR" in v: return "Acaros"
    if "AFID" in v: return "Afidos"
    if "BABOS" in v: return "Babosa"
    if "DIPTER" in v or "MOSCA" in v: return "Diptero"
    if "MINA" in v: return "Minador"
    if "MOLUS" in v or "CARAC" in v: return "Moluscos"
    if "TRIP" in v: return "Trips"

    return "OTROS"

df['blanco_norm'] = df['blanco_biologico'].apply(normalizar_blanco)

# ===============================
# 10. FILTRAR 2025
# ===============================
df_2025 = df[df['ano'] == 2025].copy()

print("\n📊 DISTRIBUCIÓN BLANCOS 2025")
print(df_2025['blanco_norm'].value_counts())

# ===============================
# 11. DONUT — DISTRIBUCIÓN GENERAL
# ===============================
dist = df_2025['blanco_norm'].value_counts().reset_index()
dist.columns = ['blanco', 'interceptaciones']

fig = px.pie(
    dist,
    names='blanco',
    values='interceptaciones',
    hole=0.45,
    title='Distribución de Interceptaciones por Blanco Biológico – 2025',
    color_discrete_sequence=PALETA_MORADO3
)

fig.update_traces(textposition='inside', textinfo='percent+label')
fig.show()


# ===============================
# 12. BARRAS APILADAS — PREDIO
# ===============================
predio_blanco = (
    df_2025
    .groupby(['predio', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    predio_blanco,
    x='predio',
    y='interceptaciones',
    color='blanco_norm',
    title='Interceptaciones 2025 por Predio y Blanco Biológico',
    text_auto=True,
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(
    barmode='stack',
    xaxis_tickangle=-45,
    plot_bgcolor='white'
)

fig.show()


# ===============================
# 13. BARRAS APILADAS — POSCOSECHA
# ===============================
pos_blanco = (
    df_2025
    .groupby(['poscosecha_proceso', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    pos_blanco,
    x='poscosecha_proceso',
    y='interceptaciones',
    color='blanco_norm',
    title='Interceptaciones 2025 por Poscosecha y Blanco Biológico',
    text_auto=True,
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(
    barmode='stack',
    xaxis_tickangle=-45,
    plot_bgcolor='white'
)

fig.show()

ORDEN_BLANCOS = ["Trips", "Afidos"]

def forzar_orden_blancos(df):
    df['blanco_norm'] = pd.Categorical(
        df['blanco_norm'],
        categories=ORDEN_BLANCOS,
        ordered=True
    )
    return df

# ===============================
# 14. BARRAS APILADAS — PAÍS
# ===============================
df_2025 = forzar_orden_blancos(df_2025)

pais_blanco = (
    df_2025
    .groupby(['pais', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    pais_blanco,
    x='pais',
    y='interceptaciones',
    color='blanco_norm',
    title='Interceptaciones 2025 por País Destino',
    text_auto=True,
    category_orders={'blanco_norm': ORDEN_BLANCOS},
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(
    barmode='stack',
    xaxis_tickangle=-45,
    plot_bgcolor='white'
)

fig.show()



# ===============================
# 15. BARRAS APILADAS — TOP 10 CLIENTES
# ===============================
top_clientes = (
    df_2025['cliente']
    .value_counts()
    .head(10)
    .index
)

cliente_blanco = (
    df_2025[df_2025['cliente'].isin(top_clientes)]
    .groupby(['cliente', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    cliente_blanco,
    x='cliente',
    y='interceptaciones',
    color='blanco_norm',
    title='Top 10 Clientes con Interceptaciones – 2025',
    text_auto=True,
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(
    barmode='stack',
    xaxis_tickangle=-45,
    plot_bgcolor='white'
)

fig.show()


print("✅ ANÁLISIS 2025 FINALIZADO – VISUALES EJECUTIVAS")

# ===============================
# 16. ANÁLISIS HISTÓRICO (2023-2025)
# ===============================
df_hist = df[
    (df['ano'].isin([2023, 2024, 2025])) &
    (df['blanco_norm'].isin(['Trips', 'Afidos']))
].copy()

# ===============================
# 17. KPI ANUAL
# ===============================
kpi_anual = (
    df_hist
    .groupby(['ano', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

print("\n📊 KPI HISTÓRICO")
print(kpi_anual)

import plotly.express as px

# ===============================
# 18. BARRAS AGRUPADAS — KPI ANUAL
# ===============================
fig = px.bar(
    kpi_anual,
    x='ano',
    y='interceptaciones',
    color='blanco_norm',
    barmode='group',
    text_auto=True,
    title='Evolución Anual de Interceptaciones – Puerto de Salida',
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(plot_bgcolor='white')
fig.show()

def forzar_orden_blancos(df):
    df['blanco_norm'] = pd.Categorical(
        df['blanco_norm'],
        categories=ORDEN_BLANCOS,
        ordered=True
    )
    return df
# ===============================
# 19. BARRAS APILADAS — PREDIOS REINCIDENTES
# ===============================
df_hist = forzar_orden_blancos(df_hist)

predios_hist = (
    df_hist
    .groupby(['predio', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

top_predios = (
    predios_hist
    .groupby('predio')['interceptaciones']
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .index
)

predios_hist = predios_hist[predios_hist['predio'].isin(top_predios)]

fig = px.bar(
    predios_hist,
    x='predio',
    y='interceptaciones',
    color='blanco_norm',
    text_auto=True,
    title='Top 10 Predios Reincidentes – Interceptaciones Históricas',
    category_orders={'blanco_norm': ORDEN_BLANCOS},
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(
    barmode='stack',
    xaxis_tickangle=-45,
    plot_bgcolor='white'
)

fig.show()


# ===============================
# 20. BARRAS APILADAS — CLIENTES REINCIDENTES
# ===============================

clientes_hist = (
    df_hist
    .groupby(['cliente', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

top_clientes = (
    clientes_hist
    .groupby('cliente')['interceptaciones']
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .index
)

clientes_hist = clientes_hist[clientes_hist['cliente'].isin(top_clientes)]

fig = px.bar(
    clientes_hist,
    x='cliente',
    y='interceptaciones',
    color='blanco_norm',
    text_auto=True,
    title='Top 10 Clientes con Interceptaciones – Histórico',
    color_discrete_map=PALETA_MORADO
)

fig.update_layout(
    barmode='stack',
    xaxis_tickangle=-45,
    plot_bgcolor='white'
)

fig.show()

# ===============================
# 21. VARIACIÓN INTERANUAL (%)
# ===============================
pivot_yoy = kpi_anual.pivot(
    index='blanco_norm',
    columns='ano',
    values='interceptaciones'
).fillna(0)

pivot_yoy['var_23_24_%'] = (
    (pivot_yoy[2024] - pivot_yoy[2023]) / pivot_yoy[2023].replace(0, 1) * 100
)

pivot_yoy['var_24_25_%'] = (
    (pivot_yoy[2025] - pivot_yoy[2024]) / pivot_yoy[2024].replace(0, 1) * 100
)

print("\n📉 VARIACIÓN INTERANUAL (%)")
print(pivot_yoy.round(1))

# ===============================
# 22. ANÁLISIS SEMÁFORO SANITARIO 2025
# ===============================
riesgo_predio = (
    df_2025
    .groupby(['predio', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

# Solo Trips y Afidos
riesgo_predio = riesgo_predio[
    riesgo_predio['blanco_norm'].isin(['Trips', 'Afidos'])
]

# Pivot
matriz = riesgo_predio.pivot_table(
    index='predio',
    columns='blanco_norm',
    values='interceptaciones',
    fill_value=0
)

# Total
matriz['Total'] = matriz.sum(axis=1)

# Top predios
matriz = matriz.sort_values('Total', ascending=False).head(15)
matriz = matriz.astype(int)

plt.figure(figsize=(8, 10))

sns.heatmap(
    matriz[['Trips', 'Afidos']],
    annot=True,
    fmt='d',
    cmap='Reds',
    linewidths=0.6,
    cbar_kws={'label': 'Número de interceptaciones'}
)

plt.title('Matriz de Riesgo Sanitario por Predio – Puerto de Salida 2025')
plt.xlabel('Blanco Biológico')
plt.ylabel('Predio')
plt.tight_layout()
plt.show()

# ===============================
# IMPACTO EN TALLOS – 2025
# ===============================

TOTAL_EXPORTADOS_2025 = 22433766

tallos_perdidos_2025 = (
    df_2025['total_tallos_rechazados']
    .sum()
)

porcentaje_perdida = (tallos_perdidos_2025 / TOTAL_EXPORTADOS_2025) * 100

impacto = pd.DataFrame({
    'categoria': ['Exportados', 'Perdidos por Interceptaciones'],
    'tallos': [TOTAL_EXPORTADOS_2025, tallos_perdidos_2025]
})

print("📉 IMPACTO PRODUCTIVO 2025")
print(f"Total exportados: {TOTAL_EXPORTADOS_2025:,}")
print(f"Tallos perdidos: {tallos_perdidos_2025:,}")
print(f"Pérdida porcentual: {porcentaje_perdida:.4f}%")
plt.figure(figsize=(8, 6))

sns.barplot(
    data=impacto,
    x='categoria',
    y='tallos',
    palette=['#5E2B97', '#B39DDB']  # morado empresa
)

plt.title('Impacto de Interceptaciones en Tallos – Puerto de Salida 2025')
plt.ylabel('Número de Tallos')
plt.xlabel('')
plt.ticklabel_format(style='plain', axis='y')

# Etiquetas
for index, row in impacto.iterrows():
    plt.text(
        index,
        row['tallos'],
        f"{row['tallos']:,}",
        ha='center',
        va='bottom',
        fontweight='bold'
    )

plt.tight_layout()
plt.show()
