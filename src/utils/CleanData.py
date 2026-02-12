# Limpieza
import re
import nltk
nltk.download('stopwords')

from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import gc
import pandas as pd
import warnings

# ------------------------------------------------------------------------------------------------------------------------------------------
# Limpieza general
# ------------------------------------------------------------------------------------------------------------------------------------------

def cast_columns(dataframe, columns_to_cast):    
    for col in columns_to_cast:
        dataframe[col] = dataframe[col].astype(str)
    return dataframe

def clean_date(date):
    return date.replace('T00:00:00.000', '').replace(' 12:00:00 AM', '')

def is_valid_date(date_str):
    try:
        date = pd.to_datetime(date_str)
        # Verificar si la fecha está dentro del rango permitido por pandas
        if pd.Timestamp.min <= date <= pd.Timestamp.max:
            return True
    except (ValueError, pd.errors.OutOfBoundsDatetime):
        return False
    return False

def convert_date_format(row, date_str, col_name, id_column_name):
    if not is_valid_date(date_str):
        contract_id = row[id_column_name]  # Obtener el ID del contrato
        # Generar un warning cuando se encuentra una fecha fuera del rango
        warnings.warn(f"Fecha inválida encontrada en la columna '{col_name}' para el contrato ID '{contract_id}': {date_str}. Reemplazando con '1900-01-01'.", UserWarning)
        return '1900-01-01'  # Reemplazar con un valor por defecto si la fecha no es válida
    try:
        # Intentar convertirlo en formato m/d/Y
        return pd.to_datetime(date_str, format='%m/%d/%Y').strftime('%Y-%m-%d')
    except ValueError:
        # Si falla, asumir que ya está en formato Y-m-d
        return pd.to_datetime(date_str).strftime('%Y-%m-%d')

def clean_date_columns(dataframe, date_columns):
    # Definir el nombre de la columna de ID
    id_column_name = 'id_adjudicacion' if 'id_adjudicacion' in dataframe.columns else 'id_contrato'
    
    for col in date_columns:
        dataframe.loc[dataframe[col] == 'nan', col] = '1900-01-01'
        dataframe.loc[(dataframe[col].isna()) | (dataframe[col] == ''), col] = '1900-01-01'
        dataframe[col] = dataframe[col].apply(clean_date)
        # Aplicar la función a la columna, pasando la fila completa, nombre de la columna, y el ID
        dataframe[col] = dataframe.apply(lambda row: convert_date_format(row, row[col], col, id_column_name), axis=1)
        dataframe[col] = pd.to_datetime(dataframe[col], errors='coerce')  # Coerce para manejar fechas inválidas
    return dataframe

def clean_url(url):
    url = url.replace("{ 'url ' : '", "").replace(" ? numconstancia= ' }", "").replace(' : ', ':')
    url = url.replace("{'url': '", "").replace("'}", "")
    return url

def clean_url_columns(dataframe, url_columns):
    for col in url_columns:
        dataframe[col] = dataframe[col].apply(clean_url)
    return dataframe

stop_words = set(stopwords.words("spanish"))

def remove_extra_punct(text):
    
    text = text.lower()
    text = re.sub(r'(?::|;|=)(?:-)?(?:\)|\(|D|P)', "", text)
    text = re.sub(r'[\\!\\"\\#\\$\\%\\&\\\'\\(\\)\\*\\+\\,\\-\\.\\/\\:\\;\\<\\=\\>\\?\\@\\[\\\\\\]\\^_\\`\\{\\|\\}\\~]', "", text)
    text = re.sub(r'\#\.', '', text)
    text = re.sub(r'\n', ' ', text)
    text = re.sub(r'  ', ' ', text)
    text = re.sub(r'´', '',text)
    text = re.sub(r',', '',text)
    text = re.sub(r'\-', '', text)
    text = re.sub(r'�', '', text)
    text = re.sub(r'á', 'a', text)
    text = re.sub(r'é', 'e', text)
    text = re.sub(r'í', 'i', text)
    text = re.sub(r'ó', 'o', text)
    text = re.sub(r'ú', 'u', text)
    text = re.sub(r'ò', 'o', text)
    text = re.sub(r'à', 'a', text)
    text = re.sub(r'è', 'e', text)
    text = re.sub(r'ì', 'i', text)
    text = re.sub(r'ù', 'u', text)
    text = re.sub("\d+", ' ', text)
    text = re.sub("\\s+", ' ', text)
    
    tokens = word_tokenize(text)
    tokens = [w for w in tokens if not w in stop_words]
    tokens = " ".join(tokens)
    
    return tokens

def clean_text_columns(dataframe, text_columns, batch_size=10000):
    num_rows = len(dataframe)
    
    for start in range(0, num_rows, batch_size):
        end = min(start + batch_size, num_rows)
        batch_df = dataframe.iloc[start:end].copy()

        for col in text_columns:
            batch_df[col] = batch_df[col].str.lower()

            # Aplicar las reglas especiales antes de la limpieza
            batch_df.loc[batch_df[col] == 'definido', col] = 'nodefinido'
            batch_df.loc[batch_df[col] == 'sin descripcion', col] = 'sindescripcion'
            batch_df.loc[batch_df[col] == 'no definido', col] = 'nodefinido'
            batch_df.loc[batch_df[col] == 'no aplica', col] = 'noaplica'

            # Crear un diccionario para almacenar los valores únicos en el lote
            unique_values = batch_df[col].unique()
            clean_value_dict = {val: remove_extra_punct(val) for val in unique_values}

            # Reasignar valores limpios a la columna
            batch_df[col] = batch_df[col].map(clean_value_dict)

            # Aplicar las reglas especiales nuevamente
            batch_df.loc[batch_df[col] == 'nodefinido', col] = 'no definido'
            batch_df.loc[batch_df[col] == 'noaplica', col] = 'no aplica'
            batch_df.loc[batch_df[col] == 'sindescripcion', col] = 'sin descripcion'

            # Liberar memoria del diccionario
            del clean_value_dict
            gc.collect()

        # Guardar o retornar el resultado en lugar de modificar todo el dataframe de una vez
        dataframe.iloc[start:end] = batch_df

        # Liberar memoria del batch
        del batch_df
        gc.collect()
    
    return dataframe

def encode_categorical_columns(dataframe, categorical_columns):
    for col in categorical_columns:
        dataframe[col] = dataframe[col].str.lower()
        dataframe.loc[dataframe[col] == 'nan', col] = -1
        dataframe.loc[dataframe[col] == 'si', col] = 1
        dataframe.loc[dataframe[col] == 'no', col] = 0
        dataframe.loc[dataframe[col] == 'no definido', col] = -1
        dataframe.loc[dataframe[col] == 'válido', col] = 1
        dataframe.loc[dataframe[col] == 'no válido', col] = 0
        dataframe.loc[dataframe[col] == 'no d', col] = -1
        dataframe[col] = dataframe[col].astype(int)
    return dataframe

def clean_s1_contracts(data: pd.DataFrame, year: int):
    date_columns = ['fecha_de_cargue_en_el_secop', 'fecha_de_firma_del_contrato', 'fecha_ini_ejec_contrato', 'fecha_fin_ejec_contrato', 
                    'fecha_liquidacion', 'ultima_actualizacion']
    url_column = ['ruta_proceso_en_secop_i']
    text_columns = ['causal_de_otras_formas_de', 'detalle_del_objeto_a_contratar', 'compromiso_presupuestal', 'objeto_del_contrato_a_la', 
                    'proponentes_seleccionados', 'calificacion_definitiva', 'posicion_rubro', 'nombre_rubro', 'pilar_acuerdo_paz', 
                    'punto_acuerdo_paz', 'cumpledecreto248', 'incluyebienesdecreto248']
    categorical_columns = ['cumple_sentencia_t302']
    columns_to_cast = []; columns_to_cast.extend(text_columns); columns_to_cast.extend(date_columns); columns_to_cast.extend(url_column);
    columns_to_cast.extend(categorical_columns)

    updated_data = cast_columns(data, columns_to_cast)
    updated_data = clean_date_columns(updated_data, date_columns)
    updated_data = clean_url_columns(updated_data, url_column)
    updated_data = clean_text_columns(updated_data, text_columns)
    updated_data = encode_categorical_columns(updated_data, categorical_columns)
    updated_data.to_csv(f'../data/silver/S1/{year}.csv', index=False, sep=';')
    return updated_data

def clean_s2_contracts(data: pd.DataFrame, year: int):
    date_columns = ['fecha_de_firma', 'fecha_de_inicio_del_contrato', 'fecha_de_fin_del_contrato', 'ultima_actualizacion', 
                    'fecha_de_inicio_de_ejecucion', 'fecha_de_fin_de_ejecucion', 'fecha_de_notificaci_n_de_prorrogaci_n', 
                    'fecha_inicio_liquidacion', 'fecha_fin_liquidacion']
    url_column = ['urlproceso']
    text_columns = ['descripcion_del_proceso', 'condiciones_de_entrega', 'habilita_pago_adelantado', 'origen_de_los_recursos', 
                    'destino_gasto', 'objeto_del_contrato']
    categorical_columns = ['liquidaci_n', 'obligaci_n_ambiental', 'obligaciones_postconsumo', 'reversion', 'espostconflicto',
                           'el_contrato_puede_ser_prorrogado']
    columns_to_cast = []; columns_to_cast.extend(text_columns); columns_to_cast.extend(date_columns); columns_to_cast.extend(url_column)
    columns_to_cast.extend(categorical_columns)

    updated_data = cast_columns(data, columns_to_cast)
    updated_data = clean_date_columns(updated_data, date_columns)
    updated_data = clean_url_columns(updated_data, url_column)
    updated_data = clean_text_columns(updated_data, text_columns)
    updated_data = encode_categorical_columns(updated_data, categorical_columns)
    updated_data = updated_data[[col for col in updated_data if 'Unnamed' not in col]]
    updated_data.to_csv(f'../data/silver/S2/{year}.csv', index=False, sep=';')
    return updated_data

def clean_numeric_columns(dataframe, numeric_columns):
    for col in numeric_columns:
        # Reemplazos iniciales
        dataframe[col] = dataframe[col].astype(str)
        dataframe[col] = dataframe[col].str.lower()
        dataframe[col] = dataframe[col].str.replace('v1.', '')
        dataframe[col] = dataframe[col].str.replace('-', '')
        dataframe[col] = dataframe[col].str.replace('.', '')
        dataframe[col] = dataframe[col].str.replace(' ', '')
    
         # Aplicar regex para extraer solo los números, eliminando cualquier otro carácter
        dataframe[col] = dataframe[col].apply(lambda x: ''.join(re.findall(r'\d+', str(x))))
        
        # Convertir valores vacíos resultantes a '0'
        dataframe[col] = dataframe[col].replace('', '0')
        dataframe[col] = dataframe[col].astype(str)
    return dataframe

# ------------------------------------------------------------------------------------------------------------------------------------------
# Limpieza de los maestros
# ------------------------------------------------------------------------------------------------------------------------------------------

def create_index_data(table_name: str, update_historical: bool = False, specific_columns=None, data: pd.DataFrame = None):
    if update_historical:
        if data is None:
            data = pd.read_csv(f'../data/bronze/{table_name}.csv', sep=';')
        columns = [col for col in data.columns if col != 'id']
        for col in columns:
            data[col] = data[col].astype(str)
            data[col] = data[col].str.lower()
        if specific_columns:
            index_data = clean_text_columns(data, specific_columns)
        else:
            index_data = clean_text_columns(data, columns)
        temporal_data = index_data.drop_duplicates(subset=columns)
        temporal_data.loc[:, 'id'] = range(1, len(temporal_data) + 1)
        index_data = pd.merge(left=index_data, right=temporal_data, how='inner', on=columns)
        index_data = index_data.rename(columns={'id_x': 'id_bronze', 'id_y': 'id_silver'})
        id_columns = ['id_bronze', 'id_silver']; id_columns.extend(columns)
        index_data[['id_bronze', 'id_silver']].to_csv(f'../data/bronze/indices/{table_name}.csv', sep=';', index=False)
        return index_data

    index_data = pd.read_csv(f'../data/bronze/indices/{table_name}.csv', sep=';')
    return index_data

def create_silver_data(table_name: str, update_historical: bool = False):
    index_data = create_index_data(table_name, update_historical)
    silver_data = index_data.drop(columns=['id_bronze'])
    silver_data = silver_data.rename(columns={'id_silver': 'id'})
    silver_data = silver_data.drop_duplicates()
    silver_data.to_csv(f'../data/silver/{table_name}.csv', sep=';', index=False)
    return index_data[['id_bronze', 'id_silver']]

def manage_new_ids(data: pd.DataFrame, table_name: str, id_column_name: str, update_historical: bool = False):
    index_data = create_silver_data(table_name, update_historical)
    updated_data = pd.merge(left=data, right=index_data, how='inner', left_on=id_column_name, right_on='id_bronze')
    if 'id_bronze' in updated_data.columns:
        columns = [id_column_name, 'id_bronze']
    else:
        columns = [id_column_name]
    updated_data = updated_data.drop(columns=columns)
    updated_data = updated_data.rename(columns={'id_silver': id_column_name})
    return updated_data

def update_id_columns(data: pd.DataFrame, update_historical: bool = False):
    if 'id_modalidad' in data.columns: updated_data = manage_new_ids(data, 'modalidad', 'id_modalidad', update_historical)
    if 'id_tipo_contrato' in data.columns: updated_data = manage_new_ids(data, 'tipo_contrato', 'id_tipo_contrato', update_historical)
    if 'id_estado_contrato' in data.columns: updated_data = manage_new_ids(data, 'estado_contrato', 'id_estado_contrato', update_historical)
    if 'id_ordenador_gasto' in data.columns: updated_data = manage_new_ids(data, 'ordenador_gasto', 'id_ordenador_gasto', update_historical)
    if 'id_supervisor' in data.columns: updated_data = manage_new_ids(data, 'supervisor', 'id_supervisor', update_historical)
    if 'id_ordenador_pago' in data.columns: updated_data = manage_new_ids(data, 'ordenador_pago', 'id_ordenador_pago', update_historical)
    return updated_data

def manage_type_columns_id(data: pd.DataFrame, columns_id: list):
    for column in columns_id:
        data[column] = data[column].astype(int)
    return data

def update_legal_reps(update_historical: bool = False):
    if update_historical:
        data = pd.read_csv(f'../data/bronze/representante_legal.csv', sep=';')
        updated_data = manage_new_ids(data, 'tipo_identificacion', 'id_tipo_documento', update_historical)
        updated_data = manage_new_ids(updated_data, 'sexo', 'id_sexo', update_historical)
        updated_data = clean_numeric_columns(updated_data, ['identificacion'])
        updated_data.loc[updated_data['id_nacionalidad'].isna(), 'id_nacionalidad'] = -1
        id_columns = ['id_tipo_documento', 'id_nacionalidad', 'id_sexo']
        updated_data = manage_type_columns_id(updated_data, id_columns)
        updated_data.loc[updated_data['domicilio'].isna(), 'domicilio'] = 'no definido'
        updated_data = create_index_data('representante_legal', update_historical, ['nombre'], updated_data)
        index_data = updated_data[['id_bronze', 'id_silver']]
        index_data.to_csv('../data/bronze/indices/representante_legal.csv', sep=';', index=False)
        updated_data = updated_data.drop(columns=['id_bronze'])
        updated_data = updated_data.drop_duplicates(subset=updated_data.columns)
        updated_data = updated_data.rename(columns={'id_silver': 'id'})
        updated_data.to_csv('../data/silver/representante_legal.csv', sep=';', index=False)
    else:
        pass

def update_locations(update_historical: bool = False):
    if update_historical:
        columns = ['departamento', 'municipio']
        data = create_index_data('ubicacion', update_historical, columns)
        for col in columns:
            data[col] = data[col].str.replace('nan', 'no definido')

        index_ubicacion = data[['id_bronze', 'id_silver']]
        index_ubicacion.to_csv('../data/bronze/indices/ubicacion.csv', sep=';', index=False)
        updated_data = data.drop(columns=['id_bronze'])
        updated_data = updated_data.drop_duplicates(subset=updated_data.columns)
        updated_data = updated_data.rename(columns={'id_silver': 'id'})
        updated_data.to_csv('../data/silver/ubicacion.csv', sep=';', index=False)
    else:
        pass

def update_providers(update_historical: bool = False):
    if update_historical:
        data = pd.read_csv('../data/bronze/proveedor.csv', sep=';')
        updated_data = manage_new_ids(data, 'tipo_identificacion', 'id_tipo_documento', update_historical)
        updated_data = cast_columns(updated_data, ['es_grupo', 'es_pyme'])
        updated_data = encode_categorical_columns(updated_data, ['es_grupo', 'es_pyme'])
        updated_data = clean_numeric_columns(updated_data, ['identificacion'])
        updated_data.loc[updated_data['codigo'].isna(), 'codigo'] = -1

        index_rep_lega = pd.read_csv('../data/bronze/indices/representante_legal.csv', sep=';')
        updated_data = pd.merge(left=updated_data, right=index_rep_lega, how='left', left_on='id_representante_legal', right_on='id_bronze')
        updated_data = updated_data.drop(columns=['id_bronze', 'id_representante_legal'])
        updated_data = updated_data.rename(columns={'id_silver': 'id_representante_legal'})

        index_ubicacion = pd.read_csv('../data/bronze/indices/ubicacion.csv', sep=';')
        updated_data = pd.merge(left=updated_data, right=index_ubicacion, how='left', left_on='id_municipio', right_on='id_bronze', indicator=True)
        updated_data.loc[updated_data['_merge'] == 'left_only', 'id_silver'] = -1
        updated_data = updated_data.drop(columns=['id_bronze', 'id_municipio', '_merge'])
        updated_data = updated_data.rename(columns={'id_silver': 'id_municipio'})

        columns_id = ['codigo', 'es_grupo', 'es_pyme', 'id_tipo_documento', 'id_representante_legal', 'id_municipio']
        updated_data = manage_type_columns_id(updated_data, columns_id)

        updated_data = create_index_data('proveedor', update_historical, ['nombre'], updated_data)

        index_data = updated_data[['id_bronze', 'id_silver']]
        index_data.to_csv('../data/bronze/indices/proveedor.csv', sep=';', index=False)
        updated_data = updated_data.drop(columns=['id_bronze'])
        updated_data = updated_data.drop_duplicates(subset=updated_data.columns)
        updated_data = updated_data.rename(columns={'id_silver': 'id'})
        updated_data.to_csv('../data/silver/proveedor.csv', sep=';', index=False)
    else:
        pass

def manage_type_masters(master_name: str, update_historical: bool = False, specific_columns = None):
    if update_historical:
        data = pd.read_csv(f'../data/bronze/{master_name}.csv', sep=';')
        columns = [col for col in data.columns if col != 'id']
        for col in columns:
            data[col] = data[col].astype(str)
        if specific_columns == None:
            updated_data = clean_text_columns(data, columns)
        else:
            updated_data = clean_text_columns(data, specific_columns)
        temporal_data = updated_data[columns].drop_duplicates(subset=columns)
        temporal_data.loc[:, 'id_silver'] = range(1, len(temporal_data) + 1)
        updated_data = pd.merge(left=updated_data, right=temporal_data, how='inner', on=columns)
        updated_data = updated_data.rename(columns={'id': 'id_bronze'})
        updated_data[['id_bronze', 'id_silver']].to_csv(f'../data/bronze/indices/{master_name}.csv', sep=';', index=False)
        updated_data = updated_data.drop(columns=['id_bronze'])
        updated_data = updated_data.rename(columns={'id_silver': 'id'})
        updated_data = updated_data.drop_duplicates(subset=updated_data.columns)
        updated_data.to_csv(f'../data/silver/{master_name}.csv', sep=';', index=False)
    else:
        pass

def update_type_masters(update_historical: bool = False):
    if update_historical:
        manage_type_masters('orden_entidad', update_historical)
        manage_type_masters('unspsc', update_historical, specific_columns=['nombre_familia', 'nombre_clase'])

        # A la fecha solo está presente en el SECOP I (11-10-2024)
        manage_type_masters('grupo', update_historical)
        manage_type_masters('moneda', update_historical)
        manage_type_masters('nivel_entidad', update_historical)
        manage_type_masters('objeto_a_contratar', update_historical)
        manage_type_masters('regimen_contratacion', update_historical)
        manage_type_masters('sub_unidad_ejecutora', update_historical)

        # A la fecha solo está presente en el SECOP II (11-10-2024)
        manage_type_masters('acuerdo', update_historical)
        manage_type_masters('bpin', update_historical)
        manage_type_masters('nacionalidad', update_historical)
        manage_type_masters('rama_entidad', update_historical)
        manage_type_masters('sector_entidad', update_historical)
    else:
        pass

def manage_new_index(data: pd.DataFrame, index_table_name: str, index_column_name: str):
    index_table = pd.read_csv(f'../data/bronze/indices/{index_table_name}.csv', sep=';')
    updated_data = pd.merge(left=data, right=index_table, how='left', left_on=index_column_name, right_on='id_bronze', indicator=True)
    updated_data.loc[updated_data['_merge'] == 'left_only', 'id_silver'] = -1
    updated_data = updated_data.drop(columns=['_merge', 'id_bronze', index_column_name])
    updated_data = updated_data.rename(columns={'id_silver': index_column_name})
    return updated_data

def update_entities(update_historical: bool = False):
    if update_historical:
        data = pd.read_csv('../data/bronze/entidad.csv', sep=';')
        updated_data = clean_numeric_columns(data, ['nit', 'codigo'])
        updated_data = manage_new_index(updated_data, 'ubicacion', 'id_ubicacion')
        updated_data = manage_new_index(updated_data, 'nivel_entidad', 'id_nivel')
        updated_data = manage_new_index(updated_data, 'orden_entidad', 'id_orden')
        updated_data = manage_new_index(updated_data, 'sub_unidad_ejecutora', 'id_sub_unidad_ejecutora')
        updated_data = manage_new_index(updated_data, 'sector_entidad', 'id_sector')
        updated_data = manage_new_index(updated_data, 'rama_entidad', 'id_rama')
        id_columns = ['id_ubicacion', 'id_nivel', 'id_orden', 'id_sub_unidad_ejecutora', 'id_sector', 'id_rama']
        updated_data = manage_type_columns_id(updated_data, id_columns)
        updated_data = create_index_data('entidad', update_historical, ['nombre'], updated_data)
        updated_data = updated_data.drop(columns=['id_bronze'])
        updated_data = updated_data.rename(columns={'id_silver': 'id'})
        updated_data = updated_data.drop_duplicates(subset=[col for col in updated_data if col != 'id'])
        updated_data.to_csv('../data/silver/entidad.csv', sep=';', index=False)
    else:
        pass

def update_db(update_historical: bool = False):
    update_legal_reps(update_historical=update_historical)
    update_locations(update_historical=update_historical)
    update_providers(update_historical=update_historical)
    update_type_masters(update_historical=update_historical)
    update_entities(update_historical=update_historical)

def generate_enriched_locations(ubdate_historical: bool = False):
    if ubdate_historical:
        data = pd.read_csv('../data/silver/ubicacion.csv', sep=';')
        # Reemplazos conocidos
        data.loc[data['departamento'].str.contains('bogota'), 'departamento'] = 'bogota . d.c .'
        data.loc[data['municipio'].str.contains('bogota'), 'municipio'] = 'bogota . d.c .'
        data.loc[data['departamento'].str.contains('narino'), 'departamento'] = 'nariño'

        maestro_municipios = pd.read_csv('../data/auxiliar/maestro_municipios.csv', sep=';')
        success_merges = []
        # Primer intento de cruce (Departamento y Municipio)
        updated_data = pd.merge(left=data, right=maestro_municipios, how='left', left_on=['departamento', 'municipio'], right_on=['DEPARTAMENTO_limpio', 'MUNICIPIO_limpio'], indicator=True)
        success_data_merge = updated_data[updated_data['_merge'] == 'both']
        success_data_merge = success_data_merge.drop(columns=['_merge'])
        success_merges.append(success_data_merge)
        error_data_merge = updated_data[updated_data['_merge'] != 'both'][['id', 'departamento', 'municipio']]

        # Segundo intento de cruce (Municipios sin homonimos)
        maestro_municipios_sin_homonimos = maestro_municipios.drop_duplicates(subset=['MUNICIPIO'])
        updated_data = pd.merge(left=error_data_merge, right=maestro_municipios_sin_homonimos, how='left', left_on=['municipio'], right_on=['MUNICIPIO_limpio'], indicator=True)
        success_data_merge = updated_data[updated_data['_merge'] == 'both']
        success_data_merge = success_data_merge.drop(columns=['_merge'])
        success_merges.append(success_data_merge)
        error_data_merge = updated_data[updated_data['_merge'] != 'both'][['id', 'departamento', 'municipio']]

        # Tercer intento de cruce (Municipios donde se reportó un departamento)
        maestro_deptos = maestro_municipios[['CODIGO_DEPARTAMENTO', 'DEPARTAMENTO', 'DEPARTAMENTO_limpio']].drop_duplicates(subset=['DEPARTAMENTO'])
        maestro_deptos['CODIGO_MUNICIPIO'] = -1
        maestro_deptos['MUNICIPIO'] = 'no definido'
        maestro_deptos['TIPO'] = 'departamento'
        maestro_deptos['CATEGORIA_MUNICIPIO'] = -1
        maestro_deptos['MUNICIPIO_limpio'] = 'no definido'
        updated_data = pd.merge(left=error_data_merge, right=maestro_deptos, how='left', left_on=['municipio'], right_on=['DEPARTAMENTO_limpio'], indicator=True)
        success_data_merge = updated_data[updated_data['_merge'] == 'both']
        success_data_merge = success_data_merge.drop(columns=['_merge'])
        success_merges.append(success_data_merge)
        error_data_merge = updated_data[updated_data['_merge'] != 'both'][['id', 'departamento', 'municipio']]

        error_data_merge['CODIGO_DEPARTAMENTO'] = -1
        error_data_merge['CODIGO_MUNICIPIO'] = -1
        error_data_merge['DEPARTAMENTO'] = error_data_merge['departamento']
        error_data_merge['MUNICIPIO'] = error_data_merge['municipio']
        error_data_merge['TIPO'] = 'no definido'
        error_data_merge['CATEGORIA_MUNICIPIO'] = -1
        error_data_merge['DEPARTAMENTO_limpio'] = error_data_merge['departamento']
        error_data_merge['MUNICIPIO_limpio'] = error_data_merge['municipio']
        success_merges.append(error_data_merge)

        updated_data = pd.concat(success_merges)
        updated_data['departamento'] = updated_data['DEPARTAMENTO'].str.title()
        updated_data['municipio'] = updated_data['MUNICIPIO'].str.title()
        updated_data = updated_data.drop(columns=['DEPARTAMENTO', 'MUNICIPIO', 'DEPARTAMENTO_limpio', 'MUNICIPIO_limpio'])
        updated_data = updated_data.rename(columns={'CODIGO_DEPARTAMENTO': 'codigo_departamento', 'CODIGO_MUNICIPIO': 'codigo_municipio', 'TIPO': 'tipo', 'CATEGORIA_MUNICIPIO': 'categoria_municipio'})

        columns = ['id', 'codigo_departamento', 'codigo_municipio', 'categoria_municipio']
        for col in columns:
            updated_data.loc[updated_data[col].isna(), col] = -1
            updated_data[col] = updated_data[col].astype(int)
        updated_data = updated_data.drop_duplicates(subset=updated_data.columns)
        updated_data.to_csv('../data/silver/ubicacion.csv', sep=';', index=False)

def update_data_base(db: pd.DataFrame, db_name: str, year: int, update_historical=False):
    db = clean_s1_contracts(db, year) if db_name.lower() == 's1' else clean_s2_contracts(db, year)
    db = update_id_columns(db, update_historical)
    generate_enriched_locations(update_historical)
    update_db(update_historical)