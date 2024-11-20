import streamlit as st
import pandas as pd
from io import BytesIO
from rapidfuzz import fuzz, process

# Cargar datos de Ramedicas desde Google Drive
@st.cache_data
def load_ramedicas_data():
    ramedicas_url = (
        "https://docs.google.com/spreadsheets/d/1Y9SgliayP_J5Vi2SdtZmGxKWwf1iY7ma/export?format=xlsx&sheet=Hoja1"
    )
    ramedicas_df = pd.read_excel(ramedicas_url, sheet_name="Hoja1")
    return ramedicas_df[['codart', 'nomart']]

# Preprocesar nombres
def preprocess_name(name):
    replacements = {
        "(": "",
        ")": "",
        "+": " ",
        "/": " ",
        "-": " ",
        ",": "",
        ";": "",
        ".": "",
        "mg": " mg",
        "ml": " ml",
        "capsula": " tableta",  # Unificar terminología
        "tablet": " tableta",
        "tableta": " tableta",
        "parches": " parche",
        "parche": " parche"
    }
    for old, new in replacements.items():
        name = name.lower().replace(old, new)
    stopwords = {"de", "el", "la", "los", "las", "un", "una", "y", "en", "por"}
    words = [word for word in name.split() if word not in stopwords]
    return " ".join(sorted(words))  # Ordenar alfabéticamente para mejorar la comparación

# Buscar la mejor coincidencia
def find_best_match_for_ramedicas(ramedicas_name, client_names_df, score_threshold=75):
    ramedicas_name_processed = preprocess_name(ramedicas_name)
    client_names_df['processed_nombre'] = client_names_df['nombre'].apply(preprocess_name)

    matches = process.extract(
        ramedicas_name_processed,
        client_names_df['processed_nombre'],
        scorer=fuzz.token_set_ratio,
        limit=5
    )

    best_match = None
    highest_score = 0

    for match, score, idx in matches:
        candidate_row = client_names_df.iloc[idx]
        if score > highest_score and score >= score_threshold:
            highest_score = score
            best_match = {
                'nombre_ramedicas': ramedicas_name,
                'nombre_cliente': candidate_row['nombre'],
                'score': score
            }
    return best_match

# Interfaz de Streamlit
st.title("Homologador de Productos - Inverso")

if st.button("Actualizar base de datos"):
    st.cache_data.clear()

uploaded_file = st.file_uploader("Sube tu archivo con los nombres de los clientes", type="xlsx")

if uploaded_file:
    client_names_df = pd.read_excel(uploaded_file)
    if 'nombre' not in client_names_df.columns:
        st.error("El archivo debe contener una columna llamada 'nombre'.")
    else:
        ramedicas_df = load_ramedicas_data()
        results = []
        for _, row in ramedicas_df.iterrows():
            match = find_best_match_for_ramedicas(row['nomart'], client_names_df)
            if match:
                match['codart'] = row['codart']
                results.append(match)
            else:
                results.append({
                    'nombre_ramedicas': row['nomart'],
                    'nombre_cliente': None,
                    'codart': row['codart'],
                    'score': 0
                })

        results_df = pd.DataFrame(results)
        st.write("Resultados de homologación inversa:")
        st.dataframe(results_df)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Homologación Inversa")
            return output.getvalue()

        st.download_button(
            label="Descargar archivo con resultados",
            data=to_excel(results_df),
            file_name="homologacion_productos_inversa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
