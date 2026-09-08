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
# COORDENADAS OFICIALES DE CABECERAS MUNICIPALES
# (idénticas a Cabeceras_Municipales.csv: se conserva la mayúscula inicial,
# las tildes y la ñ del español; este es el nombre que debe mostrarse siempre)
# ============================================
COORDS_BASE = {
        'Medellín': (6.246631, -75.581775),
        'Abejorral': (5.789315, -75.428739),
        'Abriaquí': (6.632282, -76.064304),
        'Alejandría': (6.376061, -75.141346),
        'Amagá': (6.038708, -75.702188),
        'Amalfi': (6.909655, -75.077501),
        'Andes': (5.657194, -75.878828),
        'Angelópolis': (6.109719, -75.711389),
        'Angostura': (6.885175, -75.335116),
        'Anorí': (7.074703, -75.148355),
        'Santa Fé de Antioquia': (6.556484, -75.826648),
        'Anzá': (6.302641, -75.854442),
        'Apartadó': (7.882968, -76.625279),
        'Arboletes': (8.849317, -76.426708),
        'Argelia': (5.731474, -75.141070),
        'Armenia': (6.155667, -75.786647),
        'Barbosa': (6.439195, -75.331627),
        'Belmira': (6.606319, -75.667779),
        'Bello': (6.333587, -75.555245),
        'Betania': (5.746150, -75.976790),
        'Betulia': (6.115208, -75.984452),
        'Ciudad Bolívar': (5.850273, -76.021509),
        'Briceño': (7.112803, -75.550360),
        'Buriticá': (6.720759, -75.907000),
        'Cáceres': (7.578366, -75.352050),
        'Caicedo': (6.405607, -75.982930),
        'Caldas': (6.091077, -75.633673),
        'Campamento': (6.979771, -75.298091),
        'Cañasgordas': (6.753859, -76.028228),
        'Caracolí': (6.409829, -74.757421),
        'Caramanta': (5.548530, -75.643868),
        'Carepa': (7.755148, -76.652652),
        'El Carmen de Viboral': (6.082885, -75.333901),
        'Carolina': (6.725995, -75.283192),
        'Caucasia': (7.977278, -75.197996),
        'Chigorodó': (7.666147, -76.681531),
        'Cisneros': (6.537829, -75.087047),
        'Cocorná': (6.058295, -75.185483),
        'Concepción': (6.394348, -75.257587),
        'Concordia': (6.045738, -75.908448),
        'Copacabana': (6.348557, -75.509384),
        'Dabeiba': (6.998112, -76.261614),
        'Don Matías': (6.485603, -75.392630),
        'Ebéjico': (6.325615, -75.766413),
        'El Bagre': (7.597500, -74.799097),
        'Entrerríos': (6.566273, -75.517685),
        'Envigado': (6.166695, -75.582192),
        'Fredonia': (5.928039, -75.675072),
        'Frontino': (6.776066, -76.130765),
        'Giraldo': (6.680808, -75.952158),
        'Girardota': (6.379472, -75.444238),
        'Gómez Plata': (6.683269, -75.220018),
        'Granada': (6.142892, -75.184446),
        'Guadalupe': (6.815069, -75.239862),
        'Guarne': (6.277870, -75.441612),
        'Guatapé': (6.232461, -75.160041),
        'Heliconia': (6.206757, -75.734322),
        'Hispania': (5.799461, -75.906587),
        'Itagüí': (6.175079, -75.612056),
        'Ituango': (7.171629, -75.764673),
        'Jardín': (5.597542, -75.818982),
        'Jericó': (5.789748, -75.785499),
        'La Ceja': (6.028062, -75.429433),
        'La Estrella': (6.145238, -75.637708),
        'La Pintada': (5.743808, -75.607810),
        'La Unión': (5.973845, -75.360874),
        'Liborina': (6.677316, -75.812838),
        'Maceo': (6.552116, -74.787160),
        'Marinilla': (6.173995, -75.339345),
        'Montebello': (5.946313, -75.523455),
        'Murindó': (6.977710, -76.817485),
        'Mutatá': (7.242875, -76.435875),
        'Nariño': (5.610777, -75.176262),
        'Necoclí': (8.434526, -76.787271),
        'Nechí': (8.094129, -74.776470),
        'Olaya': (6.626492, -75.811773),
        'El Peñol': (6.219349, -75.242693),
        'Peque': (7.021029, -75.910357),
        'Pueblorrico': (5.791580, -75.839903),
        'Puerto Berrío': (6.487028, -74.410016),
        'Puerto Nare': (6.186025, -74.583012),
        'Puerto Triunfo': (5.871318, -74.641190),
        'Remedios': (7.029424, -74.698135),
        'El Retiro': (6.062454, -75.501301),
        'Rionegro': (6.147148, -75.377316),
        'Sabanalarga': (6.850028, -75.816645),
        'Sabaneta': (6.149903, -75.615479),
        'Salgar': (5.964198, -75.976807),
        'San Andrés de Cuerquia': (6.916676, -75.674564),
        'San Carlos': (6.187746, -74.988097),
        'San Francisco': (5.963476, -75.101562),
        'San Jerónimo': (6.448090, -75.726975),
        'San José de la Montaña': (6.850090, -75.683352),
        'San Juan de Urabá': (8.758964, -76.528570),
        'San Luis': (6.043017, -74.993619),
        'San Pedro de los Milagros': (6.460120, -75.556743),
        'San Pedro de Urabá': (8.276884, -76.380567),
        'San Rafael': (6.293759, -75.028490),
        'San Roque': (6.485939, -75.019109),
        'San Vicente Ferrer': (6.282164, -75.332616),
        'Santa Bárbara': (5.875527, -75.567351),
        'Santa Rosa de Osos': (6.643366, -75.460723),
        'Santo Domingo': (6.473032, -75.164903),
        'El Santuario': (6.136871, -75.265465),
        'Segovia': (7.079648, -74.701596),
        'Sonsón': (5.714851, -75.309596),
        'Sopetrán': (6.500745, -75.747378),
        'Támesis': (5.664645, -75.714429),
        'Tarazá': (7.580127, -75.401407),
        'Tarso': (5.864542, -75.822956),
        'Titiribí': (6.062391, -75.791887),
        'Toledo': (7.010328, -75.692281),
        'Turbo': (8.089929, -76.728858),
        'Uramita': (6.898393, -76.173284),
        'Urrao': (6.317343, -76.133951),
        'Valdivia': (7.165200, -75.439274),
        'Valparaiso': (5.614555, -75.624452),
        'Vegachí': (6.773525, -74.798714),
        'Venecia': (5.964693, -75.735544),
        'Vigía del Fuerte': (6.588164, -76.896004),
        'Yalí': (6.676554, -74.840059),
        'Yarumal': (6.963832, -75.418828),
        'Yolombó': (6.594511, -75.013385),
        'Yondó': (7.003960, -73.912445),
        'Zaragoza': (7.488583, -74.867075),
}

def _clave_normalizada(nombre):
    """Clave de comparación (mayúsculas, sin tildes) usada ÚNICAMENTE para
    emparejar nombres de municipios escritos de forma inconsistente en los
    datos fuente (mayúsculas/minúsculas variables) con su registro oficial en
    COORDS_BASE. Nunca se usa para mostrar el nombre en pantalla."""
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

# Alias para variantes del dato fuente que no coinciden con COORDS_BASE ni
# siquiera después de quitar tildes/mayúsculas (falta un espacio o una palabra).
# La clave es la variante ya pasada por _clave_normalizada; el valor es el
# nombre oficial (tal cual aparece en COORDS_BASE) al que debe apuntar.
ALIAS_MUNICIPIOS = {
    'DONMATIAS': 'Don Matías',
    'SAN VICENTE': 'San Vicente Ferrer',
}

# Índice de coordenadas por clave normalizada, conservando siempre el nombre
# y las coordenadas EXACTAS de COORDS_BASE (con tildes, ñ y mayúscula inicial)
COORDS_POR_CLAVE = {
    _clave_normalizada(nombre): {'Municipio': nombre, 'Latitud': lat, 'Longitud': lon}
    for nombre, (lat, lon) in COORDS_BASE.items()
}
for _clave_alias, _nombre_oficial in ALIAS_MUNICIPIOS.items():
    _lat, _lon = COORDS_BASE[_nombre_oficial]
    COORDS_POR_CLAVE[_clave_alias] = {'Municipio': _nombre_oficial, 'Latitud': _lat, 'Longitud': _lon}

def cargar_coordenadas(df):
    """
    Construye la tabla de coordenadas para los municipios presentes en df.
    El nombre y las coordenadas mostradas son siempre los de COORDS_BASE;
    la clave normalizada solo se usa internamente para encontrar el registro.
    """
    filas = []
    claves_vistas = set()
    for mun in df['Municipio'].dropna().unique():
        clave = _clave_normalizada(mun)
        if not clave or clave in claves_vistas:
            continue
        claves_vistas.add(clave)
        info = COORDS_POR_CLAVE.get(clave)
        if info is None:
            print(f"⚠️ Municipio sin coordenadas en COORDS_BASE: {mun}")
            filas.append({
                'Municipio': mun, 'Municipio_Norm': clave,
                'Latitud': 6.2, 'Longitud': -75.5
            })
        else:
            filas.append({
                'Municipio': info['Municipio'], 'Municipio_Norm': clave,
                'Latitud': info['Latitud'], 'Longitud': info['Longitud']
            })

    print(f"✅ Coordenadas asignadas para {len(filas)} municipios")
    return pd.DataFrame(filas)

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
# ENSO / RONI (Página 4)
# ============================================
MESES_ABREV = {
    1: 'ene', 2: 'feb', 3: 'mar', 4: 'abr', 5: 'may', 6: 'jun',
    7: 'jul', 8: 'ago', 9: 'sep', 10: 'oct', 11: 'nov', 12: 'dic'
}

ENSO_ORDEN = ['El Niño', 'La Niña', 'Neutral', 'Sin dato']
ENSO_COLORS = {
    'El Niño': '#F4A6B7',
    'La Niña': '#5B4FCF',
    'Neutral': '#B5B5B5',
    'Sin dato': '#E0E0E0',
}

# Amenazas secundarias asociadas a lluvia que se muestran en el gráfico ENSO
# vs. Amenaza secundaria (se ignoran categorías minoritarias como Socavación,
# Colapso estructural, Granizada, etc.)
AMENAZAS_LLUVIA_ORDEN = ['Avenida torrencial', 'Creciente súbita', 'Inundación', 'Movimiento en masa', 'Vendaval']

def clasificar_amenaza_lluvia(texto):
    """Normaliza variantes de mayúsculas ('Avenida Torrencial' vs 'Avenida torrencial')
    y descarta amenazas secundarias que no están en AMENAZAS_LLUVIA_ORDEN."""
    if pd.isna(texto):
        return None
    texto_lower = str(texto).strip().lower()
    for categoria in AMENAZAS_LLUVIA_ORDEN:
        if texto_lower == categoria.lower():
            return categoria
    return None

def _tabla_mensual_enso(df):
    """Un registro por (Año, Mes) con el total de emergencias de ese mes y la
    categoría ENSO_RONI vigente ese mes (el índice RONI es mensual, no varía
    entre registros del mismo mes)."""
    base = df[(df['Año'] > 0) & df['Mes'].notna()].copy()
    if len(base) == 0:
        return pd.DataFrame(columns=['Año', 'Mes', 'Total Emergencias', 'ENSO_RONI', 'Fecha_Mes'])

    total_mensual = base.groupby(['Año', 'Mes']).size().reset_index(name='Total Emergencias')
    enso_mensual = base.groupby(['Año', 'Mes'])['ENSO_RONI'].agg(
        lambda s: s.mode().iloc[0] if len(s.mode()) else 'Sin dato'
    ).reset_index()

    tabla = total_mensual.merge(enso_mensual, on=['Año', 'Mes'])
    tabla['Fecha_Mes'] = pd.to_datetime(dict(
        year=tabla['Año'].astype(int), month=tabla['Mes'].astype(int), day=1
    ))
    return tabla.sort_values('Fecha_Mes')

def crear_dashboard_figures_enso(df, subregion_seleccionada=None, slider_value=None):
    """Página ENSO/RONI: emergencias por lluvia vs. las fases del fenómeno ENSO."""
    df_filtrado = df.copy()
    if subregion_seleccionada:
        df_filtrado = df_filtrado[df_filtrado['Subregión'] == subregion_seleccionada]
    if slider_value:
        df_filtrado = df_filtrado[(df_filtrado['Año'] >= slider_value[0]) & (df_filtrado['Año'] <= slider_value[1])]

    tabla_mensual = _tabla_mensual_enso(df_filtrado)

    # Gráfico 1: barras por fase ENSO + línea de total mensual de emergencias
    if len(tabla_mensual) > 0:
        fig1 = px.bar(
            tabla_mensual, x='Fecha_Mes', y='Total Emergencias', color='ENSO_RONI',
            color_discrete_map=ENSO_COLORS, category_orders={'ENSO_RONI': ENSO_ORDEN},
        )
        fig1.add_trace(go.Scatter(
            x=tabla_mensual['Fecha_Mes'], y=tabla_mensual['Total Emergencias'],
            mode='lines', name='Total Emergencias Lluvia', line=dict(color='black', width=2)
        ))
        fig1.update_xaxes(tickformat='%b %Y', dtick='M6', tickangle=-45)
        fig1.update_layout(
            height=350, margin=dict(l=0, r=0, t=30, b=0),
            legend_title_text='ENSO', yaxis_title='Total emergencias por lluvia'
        )
    else:
        fig1 = px.bar(pd.DataFrame({'x': ['Sin datos'], 'y': [0]}), x='x', y='y')
        fig1.update_layout(height=350, margin=dict(l=0, r=0, t=30, b=0), showlegend=False)

    # Gráfico 2: promedio mensual de emergencias por mes calendario y fase ENSO
    tabla_prom = tabla_mensual[tabla_mensual['ENSO_RONI'] != 'Sin dato']
    if len(tabla_prom) > 0:
        tabla_prom = tabla_prom.groupby(['Mes', 'ENSO_RONI'])['Total Emergencias'].mean().reset_index()
        tabla_prom['Mes_Nombre'] = tabla_prom['Mes'].astype(int).map(MESES_ABREV)
        fig2 = px.bar(
            tabla_prom, x='Mes_Nombre', y='Total Emergencias', color='ENSO_RONI', barmode='group',
            color_discrete_map=ENSO_COLORS, text_auto='.0f',
            category_orders={'Mes_Nombre': list(MESES_ABREV.values()), 'ENSO_RONI': ['El Niño', 'La Niña', 'Neutral']},
        )
        fig2.update_traces(textposition='outside')
        fig2.update_layout(
            height=350, margin=dict(l=0, r=0, t=30, b=0),
            legend_title_text='ENSO', xaxis_title='Mes', yaxis_title='Promedio mensual'
        )
    else:
        fig2 = px.bar(pd.DataFrame({'x': ['Sin datos'], 'y': [0]}), x='x', y='y')
        fig2.update_layout(height=350, margin=dict(l=0, r=0, t=30, b=0), showlegend=False)

    # Gráfico 3: 100% apilado de amenaza secundaria por fase ENSO
    df_amz = df_filtrado.copy()
    df_amz['Amenaza_Cat'] = df_amz['Amenaza secundaria'].apply(clasificar_amenaza_lluvia)
    df_amz = df_amz[df_amz['Amenaza_Cat'].notna() & (df_amz['ENSO_RONI'] != 'Sin dato')]
    tabla_amz = df_amz.groupby(['ENSO_RONI', 'Amenaza_Cat']).size().reset_index(name='Total')
    tabla_amz['Porcentaje'] = tabla_amz.groupby('ENSO_RONI')['Total'].transform(lambda s: s / s.sum() * 100)

    if len(tabla_amz) > 0:
        fig3 = px.bar(
            tabla_amz, x='ENSO_RONI', y='Porcentaje', color='Amenaza_Cat', barmode='stack',
            category_orders={'ENSO_RONI': ['La Niña', 'Neutral', 'El Niño'], 'Amenaza_Cat': AMENAZAS_LLUVIA_ORDEN},
        )
        fig3.update_traces(texttemplate='%{y:.2f}%', textposition='inside')
        fig3.update_layout(
            height=350, margin=dict(l=0, r=0, t=30, b=0),
            legend_title_text='Amenaza secundaria', xaxis_title='ENSO',
            yaxis_title='Total emergencias por lluvia (%)'
        )
    else:
        fig3 = px.bar(pd.DataFrame({'x': ['Sin datos'], 'y': [0]}), x='x', y='y')
        fig3.update_layout(height=350, margin=dict(l=0, r=0, t=30, b=0), showlegend=False)

    return fig1, fig2, fig3, len(df_filtrado)

# ============================================
# INICIALIZAR APP
# ============================================
app = dash.Dash(
    __name__,
    suppress_callback_exceptions=True,
    assets_folder='assets',
    external_stylesheets=[
        'https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap'
    ],
)
app.title = "SAMA - Dashboard de Emergencias"
server = app.server

# ============================================
# LAYOUT
# ============================================
app.layout = html.Div([
    # Stores para filtros por página
    dcc.Store(id='store-filtros-todas', data={'municipio': None, 'subregion': None}),
    dcc.Store(id='store-filtros-lluvia', data={'municipio': None, 'subregion': None}),
    dcc.Store(id='store-filtros-amenaza', data={'municipio': None, 'subregion': None}),
    
    html.Div([
        html.H2("📋 Navegación"),
        html.Hr(),
        dcc.RadioItems(
            id='radio-paginas',
            className='sidebar-nav',
            options=[
                {'label': '📊 Todas las Emergencias', 'value': 'todas'},
                {'label': '🌧️ Emergencias por Lluvia', 'value': 'lluvia'},
                {'label': '📌 Amenazas Secundarias', 'value': 'amenazas'},
                {'label': '🌊 ENSO / RONI', 'value': 'enso'}
            ],
            value='todas'
        ),
        html.Hr(),
        html.Div([
            html.P("SAMA", className='sidebar-info-title'),
            html.P("Sistema de Alertas y Monitoreo de Antioquia"),
            html.P("Datos históricos de emergencias")
        ], className='sidebar-info')
    ], className='sidebar'),
    
    html.Div([
        html.Div(id='contenido-principal')
    ], className='main-content')
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
    ], className='sama-table')
    
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
    df_mapa = df_filtrado.groupby(['Municipio', 'Subregión']).size().reset_index()
    df_mapa.columns = ['Municipio', 'Subregión', 'Total Emergencias']
    df_mapa['Municipio_Norm'] = df_mapa['Municipio'].apply(_clave_normalizada)

    # Unir con las coordenadas oficiales (nombre y coordenadas exactas de COORDS_BASE)
    df_mapa = df_mapa.merge(
        df_coordenadas[['Municipio_Norm', 'Municipio', 'Latitud', 'Longitud']],
        on='Municipio_Norm', how='inner', suffixes=('', '_oficial')
    )

    # Agrupar por el nombre oficial para combinar variantes de escritura del mismo municipio
    df_mapa = df_mapa.groupby(
        ['Municipio_oficial', 'Subregión', 'Latitud', 'Longitud'], as_index=False
    )['Total Emergencias'].sum()

    if len(df_mapa) > 0:
        fig2 = px.scatter_map(
            df_mapa, lat='Latitud', lon='Longitud', color='Subregión',
            size='Total Emergencias', size_max=25, hover_name='Municipio_oficial',
            hover_data={'Total Emergencias': True, 'Subregión': True, 'Latitud': False, 'Longitud': False},
            color_discrete_sequence=px.colors.qualitative.Set3,
            zoom=6, center={"lat": 6.2, "lon": -75.5},
            map_style="open-street-map",
        )
        fig2.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0))
    else:
        fig2 = px.scatter_map(
            pd.DataFrame({'lat': [6.2], 'lon': [-75.5], 'text': ['Sin datos']}),
            lat='lat', lon='lon', text='text', zoom=6, center={"lat": 6.2, "lon": -75.5},
            map_style="open-street-map",
        )
        fig2.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0))
    
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
    ], className='sama-table')
    
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
        html.H1(titulo, className='page-title'),
        html.P(subtitulo, className='page-subtitle'),
        html.Hr(),
        
        html.Div([
            html.Div([
                html.Label("🏙️ Municipio:"),
                dcc.Dropdown(
                    id=f'dropdown-municipio-{pagina}',
                    options=[{'label': 'Todos', 'value': None}] + [{'label': m, 'value': m} for m in municipios_unicos],
                    value=None,
                    placeholder="Selecciona un municipio..."
                )
            ], className='filter-item'),
            
            html.Div([
                html.Label("🗺️ Subregión:"),
                dcc.Dropdown(
                    id=f'dropdown-subregion-{pagina}',
                    options=[{'label': 'Todos', 'value': None}] + [{'label': s, 'value': s} for s in subregiones_unicas],
                    value=None,
                    placeholder="Selecciona una subregión..."
                )
            ], className='filter-item'),
            
            html.Div([
                html.Label("📊 Rango de Años:"),
                html.Div([
                    dcc.RangeSlider(
                        id=f'slider-rango-{pagina}',
                        min=anio_min, max=anio_max, value=[anio_min, anio_max],
                        marks={anio_min: str(anio_min), anio_max: str(anio_max)},
                        step=1,
                        allowCross=False,
                        tooltip={'placement': 'top', 'always_visible': True}
                    ),
                    html.Div(id=f'texto-rango-tabla-{pagina}', className='filter-range-text')
                ])
            ], className='filter-item filter-item-range')
        ], className='filters-panel'),
        
        html.Div([
            html.Div([
                html.H3("📊 Total Emergencias"),
                html.H2(id=f'total-{pagina}', children=f"{total:,}")
            ], className='metric-card metric-danger'),
            html.Div([
                html.H3("📍 Municipio Más Afectado"),
                html.H2(id=f'municipio-{pagina}', children=municipio_top)
            ], className='metric-card metric-primary'),
            html.Div([
                html.H3("📅 Total de Registros"),
                html.H2(id=f'total-registros-{pagina}', children=f"{len(df):,}")
            ], className='metric-card metric-success')
        ], className='metrics-row'),
        html.Hr(),
        html.Div([html.Div(id=f'filtros-{pagina}', className='filtros-activos')]),
        html.Hr(),
        html.Div([
            html.Div([
                html.Div([
                    html.H4("🏙️ Top municipios con más emergencias"),
                    dcc.Graph(id=f'grafico1-{pagina}', config={'displayModeBar': False})
                ], className='chart-card'),
                html.Div([
                    html.H4("🗺️ " + ("Total Emergencias por Subregión" if tipo == 'original' else "Municipios afectados por Subregión")),
                    dcc.Graph(id=f'grafico2-{pagina}', config={'displayModeBar': False})
                ], className='chart-card'),
                html.Div([
                    html.H4("📊 " + ("Emergencias por tipo de Amenaza" if tipo == 'original' else "Emergencias por mes del año")),
                    dcc.Graph(id=f'grafico3-{pagina}', config={'displayModeBar': False})
                ], className='chart-card')
            ], className='charts-column'),
            html.Div([
                html.H4("📅 Total Emergencias por Año"),
                html.Div(id=f'tabla-{pagina}', className='table-scroll'),
                html.Div(id=f'texto-registros-{pagina}', className='table-caption')
            ], className='table-column table-card')
        ], className='content-row')
    ])


def crear_layout_pagina_enso(df):
    """Layout de la página ENSO/RONI: emergencias por lluvia vs. las fases del fenómeno ENSO."""
    anios_validos = [int(a) for a in df[df['Año'] > 0]['Año'].unique()]
    anio_min = min(anios_validos) if anios_validos else 1950
    anio_max = max(anios_validos) if anios_validos else 2025
    # Por defecto se muestran los últimos 10 años: con el histórico completo
    # (desde 1921) el Gráfico 1 queda ilegible; el usuario puede ampliar el
    # rango arrastrando el slider.
    anio_inicio_defecto = max(anio_min, anio_max - 9)
    subregiones_unicas = sorted(df['Subregión'].dropna().unique())

    return html.Div([
        html.H1("🌊 SAMA - Emergencias por Lluvia vs. ENSO (RONI)", className='page-title'),
        html.P("Relación entre las emergencias por lluvia y las fases del fenómeno ENSO (índice RONI)", className='page-subtitle'),
        html.Hr(),

        html.Div([
            html.Div([
                html.Label("📊 Rango de Años:"),
                dcc.RangeSlider(
                    id='slider-rango-enso',
                    min=anio_min, max=anio_max, value=[anio_inicio_defecto, anio_max],
                    marks={anio_min: str(anio_min), anio_max: str(anio_max)},
                    step=1, allowCross=False,
                    tooltip={'placement': 'top', 'always_visible': True}
                )
            ], className='filter-item filter-item-range'),

            html.Div([
                html.Label("🗺️ Subregión:"),
                dcc.Dropdown(
                    id='dropdown-subregion-enso',
                    options=[{'label': 'Todas', 'value': None}] + [{'label': s, 'value': s} for s in subregiones_unicas],
                    value=None,
                    placeholder="Selecciona una subregión..."
                )
            ], className='filter-item'),
        ], className='filters-panel'),

        html.Div([
            html.Div([
                html.H3("📌 Total Emergencias por Lluvia"),
                html.H2(id='total-enso', children=f"{len(df):,}")
            ], className='metric-card metric-danger'),
        ], className='metrics-row'),
        html.Hr(),

        html.Div([
            html.Div([
                html.H4("📈 Total de emergencias por mes vs. ENSO"),
                dcc.Graph(id='grafico1-enso', config={'displayModeBar': False})
            ], className='chart-card'),
        ], className='charts-full-column'),

        html.Div([
            html.Div([
                html.H4("📊 Promedio mensual de emergencias por Mes y ENSO"),
                dcc.Graph(id='grafico2-enso', config={'displayModeBar': False})
            ], className='chart-card'),
            html.Div([
                html.H4("🌧️ Total emergencias por lluvia y Amenaza secundaria vs. ENSO"),
                dcc.Graph(id='grafico3-enso', config={'displayModeBar': False})
            ], className='chart-card'),
        ], className='charts-row-2col'),
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
    elif pagina == 'enso':
        return crear_layout_pagina_enso(df_lluvia)
    else:
        grupos_disponibles = sorted(df_emergencias[df_emergencias['Grupo_Amenaza'] != 'Sin clasificar']['Grupo_Amenaza'].unique())
        return html.Div([
            html.H1("📌 Dashboard de Amenazas Secundarias", className='page-title'),
            html.P("Selecciona un grupo de amenaza secundaria para visualizar", className='page-subtitle'),
            html.Hr(),
            html.Div([
                html.Div([
                    html.H3("Selecciona un Grupo de Amenaza:"),
                    dcc.Dropdown(
                        id='dropdown-grupo-amenaza',
                        options=[{'label': g, 'value': g} for g in grupos_disponibles],
                        placeholder="Selecciona un grupo..."
                    )
                ], className='grupo-selector-inner')
            ], className='grupo-selector'),
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
        return html.Div([html.P(
            "Selecciona un grupo de amenaza para ver el dashboard",
            className='empty-state'
        )])
    
    df_grupo = df_por_grupo.get(grupo_seleccionado, pd.DataFrame())
    if len(df_grupo) == 0:
        return html.Div([html.P(
            f"No se encontraron datos para '{grupo_seleccionado}'",
            className='empty-state empty-state-danger'
        )])
    
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
        fig_mapa = px.scatter_map(
            pd.DataFrame({'lat': [6.2], 'lon': [-75.5], 'text': ['Sin datos']}),
            lat='lat', lon='lon', text='text', zoom=6, center={"lat": 6.2, "lon": -75.5},
            map_style="open-street-map",
        )
        fig_mapa.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0))
        fig_mes = px.line(pd.DataFrame({'Mes': range(1, 13), 'Total': [0]*12}), x='Mes', y='Total')
        fig_mes.update_layout(height=300, margin=dict(l=0, r=0, t=30, b=0))
        return fig_vacio, fig_mapa, fig_mes, html.Div("Selecciona un grupo"), "0 registros", "✅ Sin filtros", "0", "N/A"
    
    df_grupo = df_por_grupo.get(grupo_seleccionado, pd.DataFrame())
    if len(df_grupo) == 0:
        fig_vacio = px.bar(pd.DataFrame({'x': ['Sin datos'], 'y': [0]}), y='x', x='y')
        fig_vacio.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0), showlegend=False)
        fig_mapa = px.scatter_map(
            pd.DataFrame({'lat': [6.2], 'lon': [-75.5], 'text': ['Sin datos']}),
            lat='lat', lon='lon', text='text', zoom=6, center={"lat": 6.2, "lon": -75.5},
            map_style="open-street-map",
        )
        fig_mapa.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0))
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
# CALLBACK DE FILTROS - PÁGINA "ENSO / RONI"
# ============================================
@app.callback(
    [Output('grafico1-enso', 'figure'),
     Output('grafico2-enso', 'figure'),
     Output('grafico3-enso', 'figure'),
     Output('total-enso', 'children')],
    [Input('dropdown-subregion-enso', 'value'),
     Input('slider-rango-enso', 'value')]
)
def update_enso(subregion, slider_value):
    fig1, fig2, fig3, total = crear_dashboard_figures_enso(df_lluvia, subregion, slider_value)
    return fig1, fig2, fig3, f"{total:,}"


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
    port = int(os.environ.get('PORT', 8050))
    app.run(debug=False, host='0.0.0.0', port=port)