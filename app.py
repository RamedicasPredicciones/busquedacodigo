import streamlit as st
import pandas as pd
from io import BytesIO
from googleapiclient.discovery import build
from google.oauth2 import service_account
from rapidfuzz import fuzz, process

# Función para cargar datos desde Google Sheets usando la API de Google Sheets
def load_google_sheet_data(sheet_url):
    # Cargar las credenciales y la API de Google Sheets
    creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    service = build('sheets', 'v4', credentials=creds)

    # Obtener el ID de la hoja de cálculo desde la URL
    spreadsheet_id = sheet_url.split('/d/')[1].split('/')[0]
    
    # Leer los datos de la hoja de cálculo
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="Hoja1").execute()
    values = result.get('values', [])

    # Crear un DataFrame con los datos obtenidos
    if not values:
        st.error('No data found in the sheet.')
        return pd.DataFrame()
    
    # Convertir a DataFrame, asumiendo que las primeras filas contienen los encabezados
    df = pd.DataFrame(values[1:], columns=values[0])
    return df

# Función de preprocesamiento para CUM (convertir a minúsculas)
def preprocess_cum(cum):
    return cum.strip().lower()  # Limpiar y poner el CUM en minúsculas

# Buscar la mejor coincidencia entre el CUM y los productos de Ramedicas
def find_best_match(cum, ramedicas_df):
    # Preprocesar el CUM
    cum_processed = preprocess_cum(cum)
    
    # Verificar si el CUM está en los datos de Ramedicas
    if cum_processed in ramedicas_df['cum'].apply(preprocess_cum).values:
        exact_match = ramedicas_df[ramedicas_df['cum'].apply(preprocess_cum) == cum_processed].iloc[0]
        return {
            'cum_cliente': cum,
            'codart': exact_match['codart'],
            'nomart': exact_match['nomart'],
            'score': 100  # Puntaje 100 para coincidencia exacta
        }

    # Si no se encuentra coincidencia exacta, realizar búsqueda difusa con RapidFuzz
    matches = process.extractOne(cum_processed, ramedicas_df['cum'].apply(preprocess_cum))
    
    if matches:
        best_match = ramedicas_df.iloc[matches[2]]
        return {
            'cum_cliente': cum,
            'codart': best_match['codart'],
            'nomart': best_match['nomart'],
            'score': matches[1]  # Puntaje de la mejor coincidencia
        }

    return None  # Si no hay coincidencias

# Función para convertir el DataFrame a un archivo Excel
def to_excel(df):
    output = BytesIO()  # Usamos BytesIO para manejar el archivo en memoria
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Homologación")  # Escribir el DataFrame al archivo
    return output.getvalue()

# Interfaz de Streamlit
st.title("Homologador de Productos - CUM a Ramedicas")

# Cargar los datos de productos de Ramedicas
ramedicas_df = load_google_sheet_data("https://docs.google.com/spreadsheets/d/1Y9SgliayP_J5Vi2SdtZmGxKWwf1iY7ma/export?format=xlsx&sheet=Hoja1")

# Opción para ingresar el CUM manualmente
cum_input = st.text_input("Ingresa el CUM del cliente", "")

# Opción para subir un archivo con CUMs
uploaded_file = st.file_uploader("Sube tu archivo con los CUMs de los clientes", type="xlsx")

if cum_input:
    # Si se ingresa un CUM manualmente
    match = find_best_match(cum_input, ramedicas_df)
    if match:
        result_df = pd.DataFrame([match])
        st.write("Resultado de homologación:")
        st.dataframe(result_df)  # Mostrar el resultado en una tabla
        st.download_button(
            label="Descargar resultado como archivo Excel",
            data=to_excel(result_df),
            file_name="resultado_homologacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No se encontró un producto que coincida con el CUM ingresado.")

elif uploaded_file:
    # Si se sube un archivo con CUMs
    client_cums_df = pd.read_excel(uploaded_file)
    
    # Verificar si la columna 'cum' está presente en el archivo subido
    if 'cum' not in client_cums_df.columns:
        st.error("El archivo debe contener una columna llamada 'cum'.")
    else:
        # Lista para almacenar los resultados de la homologación
        results = []
        
        # Buscar la mejor coincidencia para cada CUM
        for cum in client_cums_df['cum']:
            match = find_best_match(cum, ramedicas_df)
            if match:
                results.append(match)
            else:
                results.append({
                    'cum_cliente': cum,
                    'codart': None,
                    'nomart': None,
                    'score': 0
                })

        # Crear un DataFrame con los resultados
        results_df = pd.DataFrame(results)
        st.write("Resultados de homologación:")
        st.dataframe(results_df)  # Mostrar los resultados en una tabla

        # Botón para descargar los resultados como un archivo Excel
        st.download_button(
            label="Descargar archivo con resultados",
            data=to_excel(results_df),
            file_name="homologacion_productos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Por favor, ingresa un CUM o sube un archivo para obtener los resultados.")
