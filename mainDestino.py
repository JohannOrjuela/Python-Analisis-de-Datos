import pandas as pd
import unicodedata
import plotly.express as px

# ===============================
# 1. ARCHIVO Y HOJA
# ===============================
archivo = "DatosDestino.xlsx"
hoja = "Base Interc."

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
    if "PRODUCTO" in fila and "PUERTO DESTINO" in fila:
        header_row = i
        break

if header_row is None:
    raise ValueError("❌ No se encontró la fila de encabezados")

print(f"✅ Encabezados encontrados en la fila {header_row}")

# ===============================
# FUNCIÓN ESTÁNDAR DE ESTILO PLOTLY
# ===============================
def estilo_grafica(fig, mostrar_leyenda=True):
    fig.update_layout(
        font=dict(size=40),
        title=dict(font=dict(size=56)),
        xaxis=dict(
            title_font=dict(size=44),
            tickfont=dict(size=36)
        ),
        yaxis=dict(
            title_font=dict(size=44),
            tickfont=dict(size=36)
        ),
        legend=dict(font=dict(size=36)),
        bargap=0.3,
        plot_bgcolor='white'
    )

    fig.update_traces(
        textfont_size=18
    )

    fig.update_layout(showlegend=mostrar_leyenda)

    return fig

# ===============================
# 4. RELEER CON HEADER CORRECTO
# ===============================
df = pd.read_excel(archivo, sheet_name=hoja, header=header_row)

# ===============================
# 5. LIMPIEZA DE COLUMNAS
# ===============================
def limpiar_columnas(df):
    cols = []
    for col in df.columns:
        col = unicodedata.normalize('NFKD', str(col))
        col = col.encode('ascii', 'ignore').decode('utf-8')
        col = col.strip().lower().replace(" ", "_").replace(".", "")
        cols.append(col)
    df.columns = cols
    return df

df = limpiar_columnas(df)

# ===============================
# 6. LIMPIEZA DE TEXTO GENERAL
# ===============================
def limpiar_texto(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    x = unicodedata.normalize('NFKD', x)
    x = x.encode('ascii', 'ignore').decode('utf-8')
    return x.title()

campos_texto = ['producto', 'puerto_destino', 'cliente', 'blanco_biolog']

for c in campos_texto:
    if c in df.columns:
        df[c] = df[c].apply(limpiar_texto)

# ===============================
# 7. NORMALIZAR BLANCO BIOLÓGICO – DESTINO
# ===============================
def normalizar_blanco_destino(valor):
    if valor is None:
        return "No especificado"

    v = str(valor).upper()

    if any(x in v for x in ["TRIP", "THRIP", "THYSAN", "THRIPIDAE"]):
        return "Thysanoptera"

    if any(x in v for x in ["AFID", "HEMIP", "COCHIN"]):
        return "Hemiptera"

    if "ACAR" in v:
        return "Acari"

    if any(x in v for x in ["BABOS", "CARAC", "MOLUS"]):
        return "Moluscos"

    if any(x in v for x in ["DIPTER", "MOSCA"]):
        return "Diptera"

    if "MINA" in v:
        return "Minador"

    if "LEPID" in v:
        return "Lepidoptera"

    if any(x in v for x in ["GRILL", "ORTHOP"]):
        return "Orthoptera"

    if any(x in v for x in ["ENTYLOMA", "HONGO"]):
        return "Hongos"

    if "POSTURA" in v:
        return "Postura Insecto"

    return "No especificado"

df['blanco_norm'] = df['blanco_biolog'].apply(normalizar_blanco_destino)

# ===============================
# 8. NORMALIZAR CLIENTES
# ===============================
mapa_clientes = {
    "Mm Bv Europa": "MM Flower BV Europe",
    "Mm Flower Bv Europe": "MM Flower BV Europe",
    "Sunburst Farms (Elite)": "Sunburst Farms",
    "Sunburst Farms Elite": "Sunburst Farms"
}

df['cliente'] = df['cliente'].replace(mapa_clientes)

df = df[~df['cliente'].isin([
    "No Identificado",
    "No Intercep.",
    "Interceptaciones Ica"
])]

# ===============================
# 9. NORMALIZAR PRODUCTO
# ===============================
def normalizar_producto(valor):
    if valor is None:
        return None
    v = str(valor)
    if "-" in v:
        v = v.split("-")[0]
    v = v.strip().title()
    if v.lower() in ["no identificado", "no especificado", "no intercep."]:
        return None
    return v

df['producto_norm'] = df['producto'].apply(normalizar_producto)
df = df[df['producto_norm'].notna()]

# ===============================
# 10. FECHAS Y AÑO
# ===============================
df['interception_date'] = pd.to_datetime(
    df.get('interception_date'),
    errors='coerce',
    dayfirst=True
)

df['ano'] = pd.to_numeric(df.get('ano'), errors='coerce')

# ===============================
# 11. VALIDACIÓN EN CONSOLA
# ===============================
ANIO_REPORTE = 2025
ANIOS_HIST = [2023, 2024, 2025]

df_2025 = df[df['ano'] == ANIO_REPORTE]

print("\n================ VALIDACIÓN DESTINO ================")
print(f"Año del reporte: {ANIO_REPORTE}")
print(f"Total interceptaciones 2025: {len(df_2025)}\n")

print("📌 Blancos biológicos 2025:")
print(df_2025['blanco_norm'].value_counts(), "\n")

print("👥 Top clientes:")
print(df_2025['cliente'].value_counts().head(10), "\n")

print("🌸 Top productos:")
print(df_2025['producto_norm'].value_counts().head(10), "\n")

print("📊 Histórico 2023–2025:")
print(
    df[df['ano'].isin(ANIOS_HIST)]
    .groupby(['ano', 'blanco_norm'])
    .size()
)
print("🌍 Top países destino 2025:")
print(
    df_2025['puerto_destino']
    .value_counts()
    .head(10),
    "\n"
)

print("===================================================\n")

# ===============================
# CONFIGURACIÓN GRÁFICAS
# ===============================
PALETA_DESTINO = [
    '#6A0DAD', '#B19CD9', '#9B59B6',
    '#D7BDE2', '#BB8FCE', '#7D3C98'
]

# ===============================
# 1. DISTRIBUCIÓN GENERAL 2025
# ===============================
dist_blancos_2025 = (
    df_2025.groupby('blanco_norm')
    .size()
    .reset_index(name='interceptaciones')
    .sort_values('interceptaciones')
)

fig = px.bar(
    dist_blancos_2025,
    x='interceptaciones',
    y='blanco_norm',
    orientation='h',
    text='interceptaciones',
    title='Distribución de Interceptaciones por Blanco Biológico – Destino 2025',
    color='blanco_norm',
    color_discrete_sequence=PALETA_DESTINO
)

estilo_grafica(fig, mostrar_leyenda=False).show()


# ===============================
# 2. EVOLUCIÓN HISTÓRICA
# ===============================
hist_blancos = (
    df[df['ano'].isin(ANIOS_HIST)]
    .groupby(['ano', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    hist_blancos,
    x='ano',
    y='interceptaciones',
    color='blanco_norm',
    barmode='stack',
    text_auto=True,
    title='Evolución Histórica de Interceptaciones – Destino (2023–2025)',
    color_discrete_sequence=PALETA_DESTINO
)

estilo_grafica(fig).show()


# ===============================
# 3. TOP PAÍSES DESTINO – 2025
# ===============================
top_paises = (
    df_2025['puerto_destino']
    .value_counts()
    .head(10)
    .index
)

pais_2025 = (
    df_2025[df_2025['puerto_destino'].isin(top_paises)]
    .groupby(['puerto_destino', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    pais_2025,
    x='puerto_destino',
    y='interceptaciones',
    color='blanco_norm',
    barmode='stack',
    text_auto=True,
    title='Interceptaciones por País Destino – Top 10 (2025)',
    color_discrete_sequence=PALETA_DESTINO
)

fig.update_layout(xaxis_tickangle=-45)
estilo_grafica(fig).show()


# ===============================
# 4. TOP CLIENTES – 2025
# ===============================
top_clientes = df_2025['cliente'].value_counts().head(10).index

clientes_2025 = (
    df_2025[df_2025['cliente'].isin(top_clientes)]
    .groupby(['cliente', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    clientes_2025,
    x='interceptaciones',
    y='cliente',
    color='blanco_norm',
    orientation='h',
    text='interceptaciones',
    title='Top 10 Clientes con Interceptaciones – Destino 2025',
    color_discrete_sequence=PALETA_DESTINO
)

estilo_grafica(fig).show()

# ===============================
# 5. TOP PRODUCTOS – 2025
# ===============================
top_productos = df_2025['producto_norm'].value_counts().head(10).index

productos_2025 = (
    df_2025[df_2025['producto_norm'].isin(top_productos)]
    .groupby(['producto_norm', 'blanco_norm'])
    .size()
    .reset_index(name='interceptaciones')
)

fig = px.bar(
    productos_2025,
    x='interceptaciones',
    y='producto_norm',
    color='blanco_norm',
    orientation='h',
    text='interceptaciones',
    title='Top 10 Productos con Interceptaciones – Destino 2025',
    color_discrete_sequence=PALETA_DESTINO
)

estilo_grafica(fig).show()


print("✅ INFORME DESTINO LISTO (DEPURADO + VALIDADO + 5 GRÁFICAS)")
