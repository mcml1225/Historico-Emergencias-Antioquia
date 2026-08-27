"""
Dashboard de SAMA - Dash
3 páginas: 1. Todas las Emergencias (gráficos originales), 2. Lluvia (mapa + líneas), 3. Amenazas (mapa + líneas)
CON FILTROS INTERACTIVOS ENTRE GRÁFICOS
"""

import dash
from dash import dcc, html, Input, Output, State, callback, ctx
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import os

# ============================================
# CARGA DE DATOS
# ============================================
def cargar_datos():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ruta = os.path.join(base_dir, 'Historico_Emergencias_Procesado_V3.xlsx')
    if not os.path.exists(ruta):
        raise FileNotFoundError(
            f"No se encontró el archivo de datos en: {ruta}\n"
            "Coloca el archivo Excel en la misma carpeta que app_dash_filtros.py"
        )
    df = pd.read_excel(ruta)
    
    anos_vacios = df['Año'].isna().sum()
    print(f"🔍 Años vacíos encontrados: {anos_vacios}")
    
    df['Año'] = df['Año'].fillna(0).astype(int)
    df['Año_Display'] = df['Año'].apply(lambda x: "Sin año" if x == 0 else str(x))
    
    print(f"✅ Total registros: {len(df):,}")
    print(f"📊 Registros sin año: {len(df[df['Año'] == 0]):,}")
    
    return df

df_emergencias = cargar_datos()

# ============================================
# FUNCIÓN PARA NORMALIZAR MUNICIPIOS
# ============================================
def normalizar_municipio(nombre):
    """Normaliza nombres de municipios: mayúsculas, sin tildes, sin espacios extras"""
    if pd.isna(nombre):
        return ''
    nombre = str(nombre).upper().strip()
    reemplazos = {
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
        'Ñ': 'N', 'Ü': 'U'
    }
    for k, v in reemplazos.items():
        nombre = nombre.replace(k, v)
    return ' '.join(nombre.split())

# ============================================
# CARGAR COORDENADAS DE MUNICIPIOS
# ============================================
def generar_coordenadas_para_municipios(municipios_list):
    """Genera coordenadas aproximadas para una lista de municipios normalizados."""
    coords_base = {
        'MEDELLIN': (6.2442, -75.5812),
        'BELLO': (6.3355, -75.5638),
        'ITAGUI': (6.1674, -75.6207),
        'ENVIGADO': (6.1710, -75.5920),
        'SABANETA': (6.1492, -75.6160),
        'LA ESTRELLA': (6.1500, -75.6300),
        'CALDAS': (6.0900, -75.6400),
        'RIONEGRO': (6.1450, -75.3800),
        'CARMEN DE VIBORAL': (6.0800, -75.3300),
        'MARINILLA': (6.1700, -75.3300),
        'SANTUARIO': (6.1300, -75.2600),
        'ABEJORRAL': (5.8800, -75.4300),
        'LA CEJA': (6.0300, -75.4200),
        'EL RETIRO': (6.0600, -75.5000),
        'SANTA BARBARA': (5.8700, -75.5700),
        'MONTEBELLO': (5.9400, -75.5200),
        'FREDONIA': (5.9200, -75.6700),
        'AMAGA': (6.0400, -75.7000),
        'TITIRIBI': (6.0600, -75.7900),
        'ANGELOPOLIS': (6.1100, -75.7100),
        'ARMENIA': (6.1600, -75.7900),
        'CARAMANTA': (5.8200, -75.6400),
        'CONCORDIA': (6.0400, -75.9000),
        'HELICONIA': (6.2100, -75.7300),
        'JERICO': (5.7900, -75.7800),
        'PUEBLORRICO': (5.7900, -75.8400),
        'SALGAR': (5.9600, -75.9600),
        'TARSO': (5.8600, -75.8200),
        'TAMESIS': (5.6600, -75.7100),
        'VALPARAISO': (5.6100, -75.6200),
        'JARDIN': (5.6000, -75.8200),
        'ANDES': (5.6600, -75.8800),
        'BETANIA': (5.7400, -75.9700),
        'BETULIA': (6.1100, -75.9800),
        'CAICEDO': (6.4000, -75.9800),
        'ANZA': (6.3000, -75.8500),
        'BURITICA': (6.7200, -75.9000),
        'CANASGORDAS': (6.7500, -76.0200),
        'DABEIBA': (7.0000, -76.1800),
        'EBEJICO': (6.3300, -75.7700),
        'FRONTINO': (6.7800, -76.1300),
        'GIRALDO': (6.6800, -75.9500),
        'LIBORINA': (6.6700, -75.8100),
        'OLAYA': (6.6200, -75.8100),
        'PEQUE': (6.9200, -76.0300),
        'SABANALARGA': (6.8500, -75.8200),
        'SAN JERONIMO': (6.4500, -75.7300),
        'SANTA FE DE ANTIOQUIA': (6.5600, -75.8300),
        'SOPETRAN': (6.5000, -75.7400),
        'URAMITA': (6.9000, -76.1700),
        'YARUMAL': (7.0000, -75.4200),
        'ANGOSTURA': (6.9000, -75.3300),
        'BELMIRA': (6.6000, -75.6600),
        'CAROLINA': (6.7200, -75.2800),
        'DONMATIAS': (6.4800, -75.3900),
        'ENTRERRIOS': (6.5700, -75.5000),
        'GOMEZ PLATA': (6.6800, -75.2200),
        'GUADALUPE': (6.8100, -75.2400),
        'ITUANGO': (7.1700, -75.7600),
        'SAN ANDRES DE CUERQUIA': (6.9000, -75.6700),
        'SAN JOSE DE LA MONTAÑA': (6.8500, -75.6800),
        'SAN PEDRO DE LOS MILAGROS': (6.4600, -75.5600),
        'SANTA ROSA DE OSOS': (6.6400, -75.4600),
        'TOLEDO': (6.9100, -75.5700),
        'VALDIVIA': (7.3100, -75.4400),
        'CAUCASIA': (7.9800, -75.1900),
        'EL BAGRE': (7.6000, -74.8000),
        'NECHI': (8.0900, -74.7800),
        'TARAZA': (7.5800, -75.4000),
        'ZARAGOZA': (7.4900, -74.8700),
        'CACERES': (7.5800, -75.3400),
        'CISNEROS': (6.5400, -75.0800),
        'SAN ROQUE': (6.4800, -75.0200),
        'SANTO DOMINGO': (6.4700, -75.1600),
        'CONCEPCION': (6.4100, -75.2500),
        'EL PEÑOL': (6.2200, -75.2400),
        'GUARNE': (6.2800, -75.4400),
        'SAN VICENTE': (6.2800, -75.3300),
        'SAN RAFAEL': (6.2900, -75.0300),
        'ALEJANDRIA': (6.3700, -75.1400),
        'GUATAPE': (6.2300, -75.1500),
        'GRANADA': (6.1300, -75.1800),
        'LA UNION': (5.9700, -75.3600),
        'SONSON': (5.7100, -75.3100),
        'AGUADAS': (5.6100, -75.4500),
        'LA DORADA': (5.4500, -74.6600),
        'PUERTO BERRÍO': (6.4900, -74.4000),
        'PUERTO NARE': (6.1900, -74.5800),
        'PUERTO TRIUNFO': (5.8700, -74.6500),
        'TURBO': (8.0900, -76.7300),
        'APARTADO': (7.8800, -76.6200),
        'ARBOLETES': (8.8500, -76.4200),
        'CAREPA': (7.7500, -76.6500),
        'CHIGORODO': (7.6700, -76.6800),
        'MURINDO': (6.9800, -76.7700),
        'NECOCLI': (8.4300, -76.7800),
        'SAN JUAN DE URABA': (8.7600, -76.5300),
        'SAN PEDRO DE URABA': (8.2800, -76.3800),
        'VIGIA DEL FUERTE': (6.5800, -76.8800),
        'MACEO': (6.5800, -74.7800),
        'COCORNA': (6.0600, -75.1800),
        'SAN CARLOS': (6.1900, -74.9900),
        'SAN FRANCISCO': (6.1200, -75.0800),
        'ARGELIA': (5.7300, -75.1400),
        'NARIÑO': (5.6100, -75.1700),
        'URRAO': (5.9800, -76.1300),
        'CARACOLI': (6.4100, -74.7600),
        'YOLOMBO': (6.6000, -75.0100),
        'CIUDAD BOLIVAR': (5.8500, -76.0200),
        'VENECIA': (5.9600, -75.7200),
        'REMEDIOS': (7.0300, -74.4600),
        'LA PINTADA': (5.7500, -75.6000),
        'SAN LUIS': (6.0400, -75.1000),
        'YALI': (6.6700, -74.8300),
        'SEGOVIA': (7.0800, -74.7000),
        'HISPANIA': (5.8000, -75.9000),
        'ANORI': (7.0700, -75.1400),
        'MUTATA': (7.2500, -76.4300),
        'CAMPAMENTO': (6.9800, -75.3000),
        'YONDO': (7.0000, -73.9000),
        'BRICEÑO': (7.1100, -75.5500),
        'VEGACHI': (6.7700, -74.8000),
        'AMALFI': (6.9100, -75.0800),
        'EL SANTUARIO': (6.1300, -75.2600),
        'COPACABANA': (6.3500, -75.5100),
        'BARBOSA': (6.4400, -75.3300),
        'GIRARDOTA': (6.3800, -75.4500),
        'ABRIAQUI': (6.6400, -76.0600),
        'SAN PEDRO': (6.4600, -75.5600),
        'SAN JOSE': (6.8500, -75.6800),
    }
    
    coords_list = []
    for mun in municipios_list:
        mun_upper = mun.upper().strip()
        if mun_upper in coords_base:
            lat, lon = coords_base[mun_upper]
        else:
            lat, lon = 6.2, -75.5
            print(f"⚠️ Municipio sin coordenadas: {mun}")
        
        coords_list.append({
            'Municipio': mun_upper,
            'Municipio_Norm': mun_upper,
            'Latitud': lat,
            'Longitud': lon
        })
    
    return coords_list

def cargar_coordenadas(df):
    """
    Carga o genera coordenadas para todos los municipios.
    Usa la columna 'Municipio' y la normaliza.
    """
    
    # Obtener municipios únicos del DataFrame usando la columna 'Municipio'
    municipios_unicos = df['Municipio'].unique()
    municipios_normalizados = [normalizar_municipio(m) for m in municipios_unicos]
    
    print(f"📊 Municipios únicos en datos: {len(municipios_unicos)}")
    
    # Verificar si existe el archivo de coordenadas
    ruta_coordenadas = 'coordenadas_municipios_antioquia.csv'
    
    if os.path.exists(ruta_coordenadas):
        df_coords = pd.read_csv(ruta_coordenadas)
        print(f"✅ Coordenadas cargadas: {len(df_coords)} registros")
        
        # Normalizar los nombres en el archivo de coordenadas
        df_coords['Municipio_Norm'] = df_coords['Municipio'].apply(normalizar_municipio)
        
        # Crear un diccionario de coordenadas únicas por municipio normalizado
        coords_dict = {}
        for _, row in df_coords.iterrows():
            nombre_norm = row['Municipio_Norm']
            if nombre_norm and nombre_norm not in coords_dict:
                coords_dict[nombre_norm] = (row['Latitud'], row['Longitud'])
        
        print(f"📊 Municipios con coordenadas únicas: {len(coords_dict)}")
        
        # Verificar qué municipios faltan
        municipios_sin_coords = []
        for mun_norm in municipios_normalizados:
            if mun_norm and mun_norm not in coords_dict:
                municipios_sin_coords.append(mun_norm)
        
        if municipios_sin_coords:
            print(f"⚠️ {len(municipios_sin_coords)} municipios sin coordenadas, generando aproximadas...")
            nuevas_coords = generar_coordenadas_para_municipios(municipios_sin_coords)
            
            # Agregar nuevas coordenadas al diccionario
            for coord in nuevas_coords:
                nombre_norm = coord['Municipio_Norm']
                if nombre_norm not in coords_dict:
                    coords_dict[nombre_norm] = (coord['Latitud'], coord['Longitud'])
            
            print(f"✅ Total municipios con coordenadas: {len(coords_dict)}")
        
        # Crear DataFrame final con coordenadas únicas
        coords_list = []
        for nombre_norm, (lat, lon) in coords_dict.items():
            coords_list.append({
                'Municipio': nombre_norm,
                'Municipio_Norm': nombre_norm,
                'Latitud': lat,
                'Longitud': lon
            })
        
        df_coords_final = pd.DataFrame(coords_list)
        df_coords_final.to_csv(ruta_coordenadas, index=False)
        print(f"✅ Archivo de coordenadas actualizado con {len(df_coords_final)} municipios")
        
        return df_coords_final
    else:
        print("⚠️ No se encontró archivo de coordenadas. Generando coordenadas para todos los municipios...")
        coords_list = generar_coordenadas_para_municipios(municipios_normalizados)
        df_coords = pd.DataFrame(coords_list)
        df_coords.to_csv(ruta_coordenadas, index=False)
        print(f"✅ Coordenadas generadas para {len(df_coords)} municipios")
        return df_coords

df_coordenadas = cargar_coordenadas(df_emergencias)

# ============================================
# FILTRAR DATOS POR LLUVIA (Página 2)
# ============================================
df_lluvia = df_emergencias[
    df_emergencias['Amenaza'].str.contains('lluvia', case=False, na=False) &
    ~df_emergencias['Amenaza'].str.contains('temporada de menos lluvias', case=False, na=False)
].copy()

# ============================================
# DEFINIR GRUPOS DE AMENAZAS SECUNDARIAS
# ============================================
GRUPOS_AMENAZAS = {
    'Sismos / Movimientos Tectónicos': {
        'patrones': ['sismo', 'tectónico', 'tectonico', 'sísmico', 'sisnico']
    },
    'Movimientos en Masa': {
        'patrones': ['movimiento en masa', 'deslizamiento', 'derrumb', 'remoción', 'remocion', 'caída de rocas', 'caida de rocas', 'avalancha']
    },
    'Erosión / Socavación': {
        'patrones': ['erosión', 'socavación', 'socavacion', 'erosivo']
    },
    'Diapirismo': {
        'patrones': ['diapirismo']
    },
    'Inundaciones / Crecidas': {
        'patrones': ['inundación', 'inundacion', 'creciente', 'crecida', 'avenida', 'desborde']
    },
    'Vientos Fuertes / Tormentas': {
        'patrones': ['vendaval', 'vientos fuertes', 'tormenta', 'huracán', 'huracan', 'ciclón', 'ciclon', 'tornado']
    },
    'Granizada': {
        'patrones': ['granizada', 'granizo']
    },
    'Temperaturas Extremas (Frío)': {
        'patrones': ['helada', 'frío', 'frio', 'baja temperatura', 'ola de frío', 'ola de frio']
    },
    'Sequía / Déficit Hídrico': {
        'patrones': ['sequía', 'sequia', 'déficit hídrico', 'deficit hidrico', 'estiaje']
    },
    'Fenómenos Marinos / Costeros': {
        'patrones': ['marejada', 'marino', 'costero', 'oleaje', 'tsunami']
    },
    'Tormentas Eléctricas': {
        'patrones': ['tormenta eléctrica', 'tormenta electrica', 'rayo', 'relámpago', 'relampago']
    },
    'Otros Fenómenos Naturales (genérico)': {
        'patrones': ['otro', 'fenómeno natural', 'fenomeno natural', 'natural', 'evento natural']
    }
}

def clasificar_amenaza_secundaria(texto):
    if pd.isna(texto) or texto == '':
        return 'Sin clasificar'
    texto_lower = texto.lower().strip()
    for grupo, info in GRUPOS_AMENAZAS.items():
        for patron in info['patrones']:
            if patron in texto_lower:
                return grupo
    return 'Sin clasificar'

df_emergencias['Grupo_Amenaza'] = df_emergencias['Amenaza secundaria'].apply(clasificar_amenaza_secundaria)

df_por_grupo = {}
for grupo in df_emergencias['Grupo_Amenaza'].unique():
    if grupo != 'Sin clasificar':
        df_por_grupo[grupo] = df_emergencias[df_emergencias['Grupo_Amenaza'] == grupo].copy()

# ============================================
# INICIALIZAR APP
# ============================================
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "SAMA - Dashboard de Emergencias"

# ============================================
# LAYOUT
# ============================================
app.layout = html.Div([
    # Stores para filtros por página
    dcc.Store(id='store-filtros-todas', data={'municipio': None, 'subregion': None}),
    dcc.Store(id='store-filtros-lluvia', data={'municipio': None, 'subregion': None}),
    dcc.Store(id='store-filtros-amenaza', data={'municipio': None, 'subregion': None}),
    
    html.Div([
        html.H2("📋 Navegación", style={'textAlign': 'center', 'color': '#2c3e50'}),
        html.Hr(),
        dcc.RadioItems(
            id='radio-paginas',
            options=[
                {'label': '📊 Todas las Emergencias', 'value': 'todas'},
                {'label': '🌧️ Emergencias por Lluvia', 'value': 'lluvia'},
                {'label': '📌 Amenazas Secundarias', 'value': 'amenazas'}
            ],
            value='todas',
            labelStyle={'display': 'block', 'padding': '10px', 'fontSize': '16px'},
            style={'padding': '10px'}
        ),
        html.Hr(),
        html.Div([
            html.P("📊 **SAMA**", style={'fontWeight': 'bold', 'marginBottom': '5px'}),
            html.P("Sistema de Alertas y Monitoreo de Antioquia", style={'fontSize': '14px', 'color': '#666'}),
            html.P("Datos históricos de emergencias", style={'fontSize': '14px', 'color': '#666'})
        ], style={'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'})
    ], style={
        'width': '20%', 'float': 'left', 'height': '100vh', 'padding': '20px',
        'backgroundColor': '#f0f2f6', 'position': 'fixed', 'top': 0, 'left': 0,
        'overflowY': 'auto', 'boxShadow': '2px 0 5px rgba(0,0,0,0.1)'
    }),
    
    html.Div([
        html.Div(id='contenido-principal')
    ], style={'marginLeft': '22%', 'padding': '20px', 'minHeight': '100vh'})
])


# ============================================
# FUNCIONES PARA CREAR FIGURAS CON FILTROS INTERACTIVOS
# ============================================

def crear_dashboard_figures_original(df, municipio_seleccionado=None, subregion_seleccionada=None, 
                                    slider_value=None):
    """Versión original: con Subregión y Amenazas (Página 1) - CON FILTROS INTERACTIVOS (SIN FILTRO AMENAZA)"""
    df_base = df.copy()
    df_filtrado = df_base.copy()
    
    # Aplicar filtros (solo municipio y subregión)
    if municipio_seleccionado:
        df_filtrado = df_filtrado[df_filtrado['Municipio'] == municipio_seleccionado]
    if subregion_seleccionada:
        df_filtrado = df_filtrado[df_filtrado['Subregión'] == subregion_seleccionada]
    
    # Gráfico 1: Top municipios
    if municipio_seleccionado:
        top_municipios = df_filtrado.groupby('Municipio').size().reset_index()
        top_municipios.columns = ['Municipio', 'Total Emergencias']
    else:
        top_municipios = df_filtrado.groupby('Municipio').size().nlargest(10).reset_index()
        top_municipios.columns = ['Municipio', 'Total Emergencias']
    
    if len(top_municipios) == 0:
        top_municipios = pd.DataFrame({'Municipio': ['Sin datos'], 'Total Emergencias': [0]})
    top_municipios = top_municipios.sort_values('Total Emergencias', ascending=True)
    
    fig1 = px.bar(
        top_municipios, y='Municipio', x='Total Emergencias', orientation='h',
        color='Total Emergencias', color_continuous_scale='Reds', text='Total Emergencias'
    )
    fig1.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0), showlegend=False,
                      yaxis=dict(categoryorder='array', categoryarray=top_municipios['Municipio'].tolist()))
    fig1.update_traces(textposition='outside')
    
    # Gráfico 2: Subregiones
    subregiones = df_filtrado.groupby('Subregión').size().reset_index()
    subregiones.columns = ['Subregión', 'Total Emergencias']
    if len(subregiones) == 0:
        subregiones = pd.DataFrame({'Subregión': ['Sin datos'], 'Total Emergencias': [0]})
    else:
        subregiones = subregiones.sort_values('Total Emergencias', ascending=True)
    
    fig2 = px.bar(
        subregiones, y='Subregión', x='Total Emergencias', orientation='h',
        color='Total Emergencias', color_continuous_scale='Blues', text='Total Emergencias'
    )
    fig2.update_layout(height=300, margin=dict(l=0, r=0, t=0, b=0), showlegend=False,
                      yaxis=dict(categoryorder='array', categoryarray=subregiones['Subregión'].tolist()))
    fig2.update_traces(textposition='outside')
    
    # Gráfico 3: Amenazas (solo visualización, no interactivo)
    tipos = df_filtrado.groupby('Amenaza').size().reset_index()
    tipos.columns = ['Amenaza', 'Total Emergencias']
    if len(tipos) == 0:
        tipos = pd.DataFrame({'Amenaza': ['Sin datos'], 'Total Emergencias': [0]})
    else:
        tipos = tipos.sort_values('Total Emergencias', ascending=True)
        if len(tipos) > 10:
            tipos = tipos.tail(10)
    
    fig3 = px.bar(
        tipos, y='Amenaza', x='Total Emergencias', orientation='h',
        color='Total Emergencias', color_continuous_scale='Greens', text='Total Emergencias'
    )
    fig3.update_layout(height=300, margin=dict(l=0, r=0, t=0, b=0), showlegend=False,
                      yaxis=dict(categoryorder='array', categoryarray=tipos['Amenaza'].tolist()))
    fig3.update_traces(textposition='outside')
    
    # Tabla de años
    df_tabla = df_filtrado.copy()
    if slider_value:
        df_tabla = df_tabla[(df_tabla['Año'] >= slider_value[0]) & (df_tabla['Año'] <= slider_value[1])]
    
    tabla_anios = df_tabla.groupby('Año_Display').size().reset_index()
    tabla_anios.columns = ['Año', 'Total Emergencias']
    if len(tabla_anios) > 0:
        def ordenar_anios(row):
            return -1 if row['Año'] == "Sin año" else int(row['Año'])
        tabla_anios['Orden'] = tabla_anios.apply(ordenar_anios, axis=1)
        tabla_anios = tabla_anios.sort_values('Orden', ascending=False).drop('Orden', axis=1)
        tabla_anios['Total Emergencias'] = tabla_anios['Total Emergencias'].apply(lambda x: f"{x:,}")
    
    tabla_html = html.Table([
        html.Thead(html.Tr([html.Th('Año'), html.Th('Total Emergencias')])),
        html.Tbody([html.Tr([html.Td(row['Año']), html.Td(row['Total Emergencias'])]) for _, row in tabla_anios.iterrows()])
    ], style={'width': '100%', 'borderCollapse': 'collapse', 'fontSize': '14px'})
    
    total = len(df_filtrado)
    municipio_top = df_filtrado.groupby('Municipio').size().idxmax() if len(df_filtrado) > 0 else "N/A"
    
    return fig1, fig2, fig3, tabla_html, total, municipio_top, len(df_tabla)


def crear_dashboard_figures_mapa(df, municipio_seleccionado=None, subregion_seleccionada=None, 
                                 slider_value=None):
    """Versión con mapa y líneas (Páginas 2 y 3) - CON FILTROS INTERACTIVOS"""
    df_base = df.copy()
    df_filtrado = df_base.copy()
    
    # Aplicar filtros
    if municipio_seleccionado:
        df_filtrado = df_filtrado[df_filtrado['Municipio'] == municipio_seleccionado]
    if subregion_seleccionada:
        df_filtrado = df_filtrado[df_filtrado['Subregión'] == subregion_seleccionada]
    
    # Gráfico 1: Top municipios
    if municipio_seleccionado:
        top_municipios = df_filtrado.groupby('Municipio').size().reset_index()
        top_municipios.columns = ['Municipio', 'Total Emergencias']
    else:
        top_municipios = df_filtrado.groupby('Municipio').size().nlargest(10).reset_index()
        top_municipios.columns = ['Municipio', 'Total Emergencias']
    
    if len(top_municipios) == 0:
        top_municipios = pd.DataFrame({'Municipio': ['Sin datos'], 'Total Emergencias': [0]})
    top_municipios = top_municipios.sort_values('Total Emergencias', ascending=True)
    
    fig1 = px.bar(
        top_municipios, y='Municipio', x='Total Emergencias', orientation='h',
        color='Total Emergencias', color_continuous_scale='Reds', text='Total Emergencias'
    )
    fig1.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0), showlegend=False,
                      yaxis=dict(categoryorder='array', categoryarray=top_municipios['Municipio'].tolist()))
    fig1.update_traces(textposition='outside')
    
    # Gráfico 2: MAPA
    def normalizar_mun(nombre):
        if pd.isna(nombre):
            return ''
        nombre = str(nombre).upper().strip()
        reemplazos = {'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U', 'Ñ': 'N', 'Ü': 'U'}
        for k, v in reemplazos.items():
            nombre = nombre.replace(k, v)
        return ' '.join(nombre.split())
    
    df_mapa = df_filtrado.groupby(['Municipio', 'Subregión']).size().reset_index()
    df_mapa.columns = ['Municipio', 'Subregión', 'Total Emergencias']
    df_mapa['Municipio_Norm'] = df_mapa['Municipio'].apply(normalizar_mun)
    
    # Unir con coordenadas
    df_coords_norm = df_coordenadas.copy()
    df_coords_norm['Municipio_Norm'] = df_coords_norm['Municipio_Norm'].apply(lambda x: x.upper().strip() if pd.notna(x) else '')
    
    df_mapa = df_mapa.merge(df_coords_norm, on='Municipio_Norm', how='inner')
    
    if len(df_mapa) > 0:
        fig2 = px.scatter_mapbox(
            df_mapa, lat='Latitud', lon='Longitud', color='Subregión',
            size='Total Emergencias', size_max=25, hover_name='Municipio_x',
            hover_data={'Total Emergencias': True, 'Subregión': True, 'Latitud': False, 'Longitud': False},
            color_discrete_sequence=px.colors.qualitative.Set3,
            zoom=6, center={"lat": 6.2, "lon": -75.5},
        )
        fig2.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0), mapbox_style="open-street-map")
    else:
        fig2 = px.scatter_mapbox(
            pd.DataFrame({'lat': [6.2], 'lon': [-75.5], 'text': ['Sin datos']}),
            lat='lat', lon='lon', text='text', zoom=6, center={"lat": 6.2, "lon": -75.5}
        )
        fig2.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0), mapbox_style="open-street-map")
    
    # Gráfico 3: LÍNEAS POR MES
    if 'Mes' in df_filtrado.columns and len(df_filtrado) > 0:
        df_mes = df_filtrado.groupby('Mes').size().reset_index()
        df_mes.columns = ['Mes', 'Total Emergencias']
        meses_completos = pd.DataFrame({'Mes': range(1, 13)})
        df_mes = meses_completos.merge(df_mes, on='Mes', how='left').fillna(0)
        
        fig3 = px.line(
            df_mes, x='Mes', y='Total Emergencias',
            markers=True, text='Total Emergencias',
        )
        fig3.update_traces(
            textposition='top center',
            line=dict(color='#2E86C1', width=3),
            marker=dict(size=10, color='#2E86C1'),
            textfont=dict(size=12, color='#2E86C1')
        )
        fig3.update_layout(
            height=300, margin=dict(l=0, r=0, t=30, b=0),
            xaxis=dict(tickmode='linear', tick0=1, dtick=1, range=[0.5, 12.5]),
            yaxis=dict(title='Total Emergencias')
        )
    else:
        fig3 = px.line(
            pd.DataFrame({'Mes': range(1, 13), 'Total Emergencias': [0]*12}),
            x='Mes', y='Total Emergencias', markers=True, text='Total Emergencias',
        )
        fig3.update_layout(height=300, margin=dict(l=0, r=0, t=30, b=0))
        fig3.update_traces(textposition='top center')
    
    # Tabla de años
    df_tabla = df_filtrado.copy()
    if slider_value:
        df_tabla = df_tabla[(df_tabla['Año'] >= slider_value[0]) & (df_tabla['Año'] <= slider_value[1])]
    
    tabla_anios = df_tabla.groupby('Año_Display').size().reset_index()
    tabla_anios.columns = ['Año', 'Total Emergencias']
    if len(tabla_anios) > 0:
        def ordenar_anios(row):
            return -1 if row['Año'] == "Sin año" else int(row['Año'])
        tabla_anios['Orden'] = tabla_anios.apply(ordenar_anios, axis=1)
        tabla_anios = tabla_anios.sort_values('Orden', ascending=False).drop('Orden', axis=1)
        tabla_anios['Total Emergencias'] = tabla_anios['Total Emergencias'].apply(lambda x: f"{x:,}")
    
    tabla_html = html.Table([
        html.Thead(html.Tr([html.Th('Año'), html.Th('Total Emergencias')])),
        html.Tbody([html.Tr([html.Td(row['Año']), html.Td(row['Total Emergencias'])]) for _, row in tabla_anios.iterrows()])
    ], style={'width': '100%', 'borderCollapse': 'collapse', 'fontSize': '14px'})
    
    total = len(df_filtrado)
    municipio_top = df_filtrado.groupby('Municipio').size().idxmax() if len(df_filtrado) > 0 else "N/A"
    
    return fig1, fig2, fig3, tabla_html, total, municipio_top, len(df_tabla)


# ============================================
# LAYOUTS DE PÁGINAS
# ============================================

def crear_layout_pagina(titulo, subtitulo, df, pagina, tipo):
    """Crea el layout de una página"""
    total = len(df)
    municipio_top = df.groupby('Municipio').size().idxmax() if len(df) > 0 else "N/A"
    anios_validos = [int(a) for a in df[df['Año'] > 0]['Año'].unique()]
    anio_min = min(anios_validos) if anios_validos else 1903
    anio_max = max(anios_validos) if anios_validos else 2026
    
    # Obtener listas únicas para filtros
    municipios_unicos = sorted(df['Municipio'].dropna().unique())
    subregiones_unicas = sorted(df['Subregión'].dropna().unique())
    
    # Para la página 1, los filtros son solo municipio y subregión
    mostrar_amenaza = (tipo == 'original' and pagina == 'todas')
    
    return html.Div([
        html.H1(titulo, style={'textAlign': 'center'}),
        html.P(subtitulo, style={'textAlign': 'center', 'color': '#666'}),
        html.Hr(),
        
        # Filtros interactivos
        html.Div([
            html.Div([
                html.Label("🏙️ Municipio:"),
                dcc.Dropdown(
                    id=f'dropdown-municipio-{pagina}',
                    options=[{'label': 'Todos', 'value': None}] + [{'label': m, 'value': m} for m in municipios_unicos],
                    value=None,
                    placeholder="Selecciona un municipio..."
                )
            ], style={'width': '30%', 'display': 'inline-block', 'padding': '10px'}),
            
            html.Div([
                html.Label("🗺️ Subregión:"),
                dcc.Dropdown(
                    id=f'dropdown-subregion-{pagina}',
                    options=[{'label': 'Todos', 'value': None}] + [{'label': s, 'value': s} for s in subregiones_unicas],
                    value=None,
                    placeholder="Selecciona una subregión..."
                )
            ], style={'width': '30%', 'display': 'inline-block', 'padding': '10px', 'marginLeft': '2%'}),
            
            html.Div([
                html.Label("📊 Rango de Años:"),
                html.Div([
                    dcc.RangeSlider(
                        id=f'slider-rango-{pagina}',
                        min=anio_min, max=anio_max, value=[anio_min, anio_max],
                        marks={str(i): str(i) for i in range(anio_min, anio_max+1, 10)},
                        step=1
                    ),
                    html.Div(id=f'texto-rango-tabla-{pagina}', style={'textAlign': 'center', 'marginTop': '10px'})
                ])
            ], style={'width': '30%', 'display': 'inline-block', 'padding': '10px', 'marginLeft': '2%'})
        ], style={'display': 'flex', 'justifyContent': 'space-between', 'padding': '20px', 'backgroundColor': '#f8f9fa', 'borderRadius': '10px', 'marginBottom': '20px'}),
        
        # Cards de métricas
        html.Div([
            html.Div([
                html.H3("📊 Total Emergencias"),
                html.H2(id=f'total-{pagina}', children=f"{total:,}", style={'color': '#FF4444'})
            ], style={'width': '25%', 'display': 'inline-block', 'textAlign': 'center'}),
            html.Div([
                html.H3("📍 Municipio Más Afectado"),
                html.H2(id=f'municipio-{pagina}', children=municipio_top, style={'color': '#4444FF'})
            ], style={'width': '25%', 'display': 'inline-block', 'textAlign': 'center', 'marginLeft': '5%'}),
            html.Div([
                html.H3("📅 Total de Registros"),
                html.H2(id=f'total-registros-{pagina}', children=f"{len(df):,}", style={'color': '#28A745'})
            ], style={'width': '25%', 'display': 'inline-block', 'textAlign': 'center', 'marginLeft': '5%'})
        ], style={'display': 'flex', 'justifyContent': 'space-around', 'padding': '20px'}),
        html.Hr(),
        html.Div([html.Div(id=f'filtros-{pagina}', style={'fontSize': '16px', 'padding': '10px'})]),
        html.Hr(),
        html.Div([
            html.Div([
                html.Div([html.H4("🏙️ Top municipios con más emergencias"), dcc.Graph(id=f'grafico1-{pagina}', config={'displayModeBar': False})], style={'marginBottom': '20px'}),
                html.Div([html.H4("🗺️ " + ("Total Emergencias por Subregión" if tipo == 'original' else "Municipios afectados por Subregión")), dcc.Graph(id=f'grafico2-{pagina}', config={'displayModeBar': False})], style={'marginBottom': '20px'}),
                html.Div([html.H4("📊 " + ("Emergencias por tipo de Amenaza" if tipo == 'original' else "Emergencias por mes del año")), dcc.Graph(id=f'grafico3-{pagina}', config={'displayModeBar': False})])
            ], style={'width': '70%', 'display': 'inline-block', 'verticalAlign': 'top', 'paddingRight': '20px'}),
            html.Div([
                html.H4("📅 Total Emergencias por Año"),
                html.Div(id=f'tabla-{pagina}', style={'maxHeight': '600px', 'overflowY': 'auto'}),
                html.Div(id=f'texto-registros-{pagina}', style={'marginTop': '10px', 'fontSize': '14px', 'color': '#666'})
            ], style={'width': '30%', 'display': 'inline-block', 'verticalAlign': 'top'})
        ], style={'display': 'flex'})
    ])


# ============================================
# CALLBACK PARA CAMBIAR DE PÁGINA
# ============================================
@app.callback(
    Output('contenido-principal', 'children'),
    Input('radio-paginas', 'value')
)
def cambiar_pagina(pagina):
    if pagina == 'todas':
        return crear_layout_pagina(
            "📊 SAMA - Dashboard de Emergencias",
            "Todas las emergencias registradas en el sistema",
            df_emergencias, 'todas', 'original'
        )
    elif pagina == 'lluvia':
        return crear_layout_pagina(
            "🌧️ SAMA - Dashboard de Emergencias por Lluvia",
            "Emergencias asociadas a eventos de lluvia",
            df_lluvia, 'lluvia', 'mapa'
        )
    else:
        grupos_disponibles = sorted(df_emergencias[df_emergencias['Grupo_Amenaza'] != 'Sin clasificar']['Grupo_Amenaza'].unique())
        return html.Div([
            html.H1("📌 Dashboard de Amenazas Secundarias", style={'textAlign': 'center'}),
            html.P("Selecciona un grupo de amenaza secundaria para visualizar", style={'textAlign': 'center', 'color': '#666'}),
            html.Hr(),
            html.Div([
                html.Div([
                    html.H3("Selecciona un Grupo de Amenaza:"),
                    dcc.Dropdown(
                        id='dropdown-grupo-amenaza',
                        options=[{'label': g, 'value': g} for g in grupos_disponibles],
                        placeholder="Selecciona un grupo..."
                    )
                ], style={'width': '50%', 'margin': '0 auto'})
            ], style={'padding': '20px'}),
            html.Hr(),
            html.Div(id='contenido-amenaza')
        ])


# ============================================
# CALLBACK PARA AMENAZAS SECUNDARIAS
# ============================================
@app.callback(
    Output('contenido-amenaza', 'children'),
    Input('dropdown-grupo-amenaza', 'value')
)
def actualizar_grupo_amenaza(grupo_seleccionado):
    if not grupo_seleccionado:
        return html.Div([html.P("Selecciona un grupo de amenaza para ver el dashboard", 
                       style={'textAlign': 'center', 'fontSize': '18px', 'color': '#666', 'padding': '40px'})])
    
    df_grupo = df_por_grupo.get(grupo_seleccionado, pd.DataFrame())
    if len(df_grupo) == 0:
        return html.Div([html.P(f"No se encontraron datos para '{grupo_seleccionado}'", 
                       style={'textAlign': 'center', 'fontSize': '18px', 'color': '#FF4444', 'padding': '40px'})])
    
    return crear_layout_pagina(
        f"📌 Dashboard de {grupo_seleccionado}",
        f"Emergencias clasificadas como '{grupo_seleccionado}'",
        df_grupo, 'amenaza', 'mapa'
    )


# ============================================
# CALLBACKS DE FILTROS INTERACTIVOS - PÁGINA "TODAS"
# ============================================

# Callback para actualizar filtros desde gráficos - Página "Todas"
@app.callback(
    [Output('store-filtros-todas', 'data'),
     Output('dropdown-municipio-todas', 'value'),
     Output('dropdown-subregion-todas', 'value')],
    [Input('grafico1-todas', 'clickData'),
     Input('grafico2-todas', 'clickData'),
     Input('dropdown-municipio-todas', 'value'),
     Input('dropdown-subregion-todas', 'value')],
    [State('store-filtros-todas', 'data')]
)
def actualizar_filtros_todas(click1, click2, municipio_drop, subregion_drop, store_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return store_data, municipio_drop, subregion_drop
    
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    # Inicializar filtros con valores actuales
    filtros = store_data.copy() if store_data else {'municipio': None, 'subregion': None}
    nuevo_municipio = municipio_drop
    nueva_subregion = subregion_drop
    
    # Procesar clicks en gráficos
    if trigger_id == 'grafico1-todas' and click1:
        point = click1['points'][0]
        if 'y' in point:
            nuevo_municipio = point['y']
            filtros['municipio'] = nuevo_municipio
    
    elif trigger_id == 'grafico2-todas' and click2:
        point = click2['points'][0]
        if 'y' in point:
            nueva_subregion = point['y']
            filtros['subregion'] = nueva_subregion
    
    # Actualizar desde dropdowns
    elif trigger_id == 'dropdown-municipio-todas':
        filtros['municipio'] = municipio_drop
    
    elif trigger_id == 'dropdown-subregion-todas':
        filtros['subregion'] = subregion_drop
    
    return filtros, nuevo_municipio, nueva_subregion


# Callback para actualizar gráficos - Página "Todas"
@app.callback(
    [Output(f'grafico1-todas', 'figure'),
     Output(f'grafico2-todas', 'figure'),
     Output(f'grafico3-todas', 'figure'),
     Output(f'tabla-todas', 'children'),
     Output(f'texto-registros-todas', 'children'),
     Output(f'filtros-todas', 'children'),
     Output(f'total-todas', 'children'),
     Output(f'municipio-todas', 'children'),
     Output(f'total-registros-todas', 'children')],
    [Input('store-filtros-todas', 'data'),
     Input(f'slider-rango-todas', 'value')]
)
def update_todas(store_filtros, slider_value):
    municipio = store_filtros.get('municipio') if store_filtros else None
    subregion = store_filtros.get('subregion') if store_filtros else None
    
    fig1, fig2, fig3, tabla, total, top, registros = crear_dashboard_figures_original(
        df_emergencias, municipio, subregion, slider_value
    )
    
    filtros_parts = []
    if municipio: filtros_parts.append(f"🏙️ {municipio}")
    if subregion: filtros_parts.append(f"🗺️ {subregion}")
    
    filtros_text = "🔍 Filtros activos: " + (" | ".join(filtros_parts) if filtros_parts else "✅ Sin filtros activos")
    
    return fig1, fig2, fig3, tabla, f"📌 {total:,} registros totales | {registros:,} en tabla", filtros_text, f"{total:,}", top, f"{total:,}"


# ============================================
# CALLBACKS DE FILTROS INTERACTIVOS - PÁGINA "LLUVIA"
# ============================================

@app.callback(
    [Output('store-filtros-lluvia', 'data'),
     Output('dropdown-municipio-lluvia', 'value'),
     Output('dropdown-subregion-lluvia', 'value')],
    [Input('grafico1-lluvia', 'clickData'),
     Input('grafico2-lluvia', 'clickData'),
     Input('grafico3-lluvia', 'clickData'),
     Input('dropdown-municipio-lluvia', 'value'),
     Input('dropdown-subregion-lluvia', 'value')],
    [State('store-filtros-lluvia', 'data')]
)
def actualizar_filtros_lluvia(click1, click2, click3, municipio_drop, subregion_drop, store_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return store_data, municipio_drop, subregion_drop
    
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    filtros = store_data.copy() if store_data else {'municipio': None, 'subregion': None}
    nuevo_municipio = municipio_drop
    nueva_subregion = subregion_drop
    
    if trigger_id == 'grafico1-lluvia' and click1:
        point = click1['points'][0]
        if 'y' in point:
            nuevo_municipio = point['y']
            filtros['municipio'] = nuevo_municipio
    
    elif trigger_id == 'grafico2-lluvia' and click2:
        point = click2['points'][0]
        if 'color' in point and point.get('color') != 'Sin datos':
            nueva_subregion = point['color']
            filtros['subregion'] = nueva_subregion
    
    elif trigger_id == 'dropdown-municipio-lluvia':
        filtros['municipio'] = municipio_drop
    
    elif trigger_id == 'dropdown-subregion-lluvia':
        filtros['subregion'] = subregion_drop
    
    return filtros, nuevo_municipio, nueva_subregion


@app.callback(
    [Output(f'grafico1-lluvia', 'figure'),
     Output(f'grafico2-lluvia', 'figure'),
     Output(f'grafico3-lluvia', 'figure'),
     Output(f'tabla-lluvia', 'children'),
     Output(f'texto-registros-lluvia', 'children'),
     Output(f'filtros-lluvia', 'children'),
     Output(f'total-lluvia', 'children'),
     Output(f'municipio-lluvia', 'children')],
    [Input('store-filtros-lluvia', 'data'),
     Input(f'slider-rango-lluvia', 'value')]
)
def update_lluvia(store_filtros, slider_value):
    municipio = store_filtros.get('municipio') if store_filtros else None
    subregion = store_filtros.get('subregion') if store_filtros else None
    
    fig1, fig2, fig3, tabla, total, top, registros = crear_dashboard_figures_mapa(
        df_lluvia, municipio, subregion, slider_value
    )
    
    filtros_parts = []
    if municipio: filtros_parts.append(f"🏙️ {municipio}")
    if subregion: filtros_parts.append(f"🗺️ {subregion}")
    
    filtros_text = "🔍 Filtros activos: " + (" | ".join(filtros_parts) if filtros_parts else "✅ Sin filtros activos")
    
    return fig1, fig2, fig3, tabla, f"📌 {total:,} registros totales | {registros:,} en tabla", filtros_text, f"{total:,}", top


# ============================================
# CALLBACKS DE FILTROS INTERACTIVOS - PÁGINA "AMENAZA"
# ============================================

@app.callback(
    [Output('store-filtros-amenaza', 'data'),
     Output('dropdown-municipio-amenaza', 'value'),
     Output('dropdown-subregion-amenaza', 'value')],
    [Input('grafico1-amenaza', 'clickData'),
     Input('grafico2-amenaza', 'clickData'),
     Input('grafico3-amenaza', 'clickData'),
     Input('dropdown-municipio-amenaza', 'value'),
     Input('dropdown-subregion-amenaza', 'value')],
    [State('store-filtros-amenaza', 'data')]
)
def actualizar_filtros_amenaza(click1, click2, click3, municipio_drop, subregion_drop, store_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return store_data, municipio_drop, subregion_drop
    
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    filtros = store_data.copy() if store_data else {'municipio': None, 'subregion': None}
    nuevo_municipio = municipio_drop
    nueva_subregion = subregion_drop
    
    if trigger_id == 'grafico1-amenaza' and click1:
        point = click1['points'][0]
        if 'y' in point:
            nuevo_municipio = point['y']
            filtros['municipio'] = nuevo_municipio
    
    elif trigger_id == 'grafico2-amenaza' and click2:
        point = click2['points'][0]
        if 'color' in point and point.get('color') != 'Sin datos':
            nueva_subregion = point['color']
            filtros['subregion'] = nueva_subregion
    
    elif trigger_id == 'dropdown-municipio-amenaza':
        filtros['municipio'] = municipio_drop
    
    elif trigger_id == 'dropdown-subregion-amenaza':
        filtros['subregion'] = subregion_drop
    
    return filtros, nuevo_municipio, nueva_subregion


@app.callback(
    [Output(f'grafico1-amenaza', 'figure'),
     Output(f'grafico2-amenaza', 'figure'),
     Output(f'grafico3-amenaza', 'figure'),
     Output(f'tabla-amenaza', 'children'),
     Output(f'texto-registros-amenaza', 'children'),
     Output(f'filtros-amenaza', 'children'),
     Output(f'total-amenaza', 'children'),
     Output(f'municipio-amenaza', 'children')],
    [Input('store-filtros-amenaza', 'data'),
     Input(f'slider-rango-amenaza', 'value'),
     Input('dropdown-grupo-amenaza', 'value')]
)
def update_amenaza(store_filtros, slider_value, grupo_seleccionado):
    if not grupo_seleccionado:
        fig_vacio = px.bar(pd.DataFrame({'x': ['Sin datos'], 'y': [0]}), y='x', x='y')
        fig_vacio.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0), showlegend=False)
        fig_mapa = px.scatter_mapbox(
            pd.DataFrame({'lat': [6.2], 'lon': [-75.5], 'text': ['Sin datos']}),
            lat='lat', lon='lon', text='text', zoom=6, center={"lat": 6.2, "lon": -75.5}
        )
        fig_mapa.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0), mapbox_style="open-street-map")
        fig_mes = px.line(pd.DataFrame({'Mes': range(1, 13), 'Total': [0]*12}), x='Mes', y='Total')
        fig_mes.update_layout(height=300, margin=dict(l=0, r=0, t=30, b=0))
        return fig_vacio, fig_mapa, fig_mes, html.Div("Selecciona un grupo"), "0 registros", "✅ Sin filtros", "0", "N/A"
    
    df_grupo = df_por_grupo.get(grupo_seleccionado, pd.DataFrame())
    if len(df_grupo) == 0:
        fig_vacio = px.bar(pd.DataFrame({'x': ['Sin datos'], 'y': [0]}), y='x', x='y')
        fig_vacio.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0), showlegend=False)
        fig_mapa = px.scatter_mapbox(
            pd.DataFrame({'lat': [6.2], 'lon': [-75.5], 'text': ['Sin datos']}),
            lat='lat', lon='lon', text='text', zoom=6, center={"lat": 6.2, "lon": -75.5}
        )
        fig_mapa.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0), mapbox_style="open-street-map")
        fig_mes = px.line(pd.DataFrame({'Mes': range(1, 13), 'Total': [0]*12}), x='Mes', y='Total')
        fig_mes.update_layout(height=300, margin=dict(l=0, r=0, t=30, b=0))
        return fig_vacio, fig_mapa, fig_mes, html.Div("No hay datos"), "0 registros", "✅ Sin filtros", "0", "N/A"
    
    municipio = store_filtros.get('municipio') if store_filtros else None
    subregion = store_filtros.get('subregion') if store_filtros else None
    
    fig1, fig2, fig3, tabla, total, top, registros = crear_dashboard_figures_mapa(
        df_grupo, municipio, subregion, slider_value
    )
    
    filtros_parts = []
    if municipio: filtros_parts.append(f"🏙️ {municipio}")
    if subregion: filtros_parts.append(f"🗺️ {subregion}")
    
    filtros_text = "🔍 Filtros activos: " + (" | ".join(filtros_parts) if filtros_parts else "✅ Sin filtros activos")
    
    return fig1, fig2, fig3, tabla, f"📌 {total:,} registros totales | {registros:,} en tabla", filtros_text, f"{total:,}", top


# ============================================
# CALLBACKS PARA TEXTO DE RANGO DE AÑOS
# ============================================
@app.callback(
    Output('texto-rango-tabla-todas', 'children'),
    Input('slider-rango-todas', 'value')
)
def update_texto_todas(value):
    return f"📅 Años en tabla: {value[0]} - {value[1]}"

@app.callback(
    Output('texto-rango-tabla-lluvia', 'children'),
    Input('slider-rango-lluvia', 'value')
)
def update_texto_lluvia(value):
    return f"📅 Años en tabla: {value[0]} - {value[1]}"

@app.callback(
    Output('texto-rango-tabla-amenaza', 'children'),
    Input('slider-rango-amenaza', 'value')
)
def update_texto_amenaza(value):
    return f"📅 Años en tabla: {value[0]} - {value[1]}"


# ============================================
# EJECUTAR APP
# ============================================
if __name__ == '__main__':
    app.run(debug=True, port=8050)