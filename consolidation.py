from functions import Functions
import io
import os
from datetime import timedelta
import pandas as pd
import shutil
import openpyxl as xl
import locale
from pandas.errors import MergeError
from datetime import datetime

# Definir región
locale.setlocale(locale.LC_ALL, 'es_CO')

class Consolidation(Functions):

    def __init__(self, credentials, start_date, end_date):
        super().__init__()
        self.nac_user = credentials['login_NACIONAL_user'] # Usuario nacional
        self.nac_pwd = credentials['login_NACIONAL_pwd'] # Contraseña nacional
        self.start_date = start_date # Fecha inicio
        self.end_date = end_date # Fecha fin
        self.conn = self.connect_odbc('AS400', maq = 'NACIONAL', user = self.nac_user, pwd = self.nac_pwd)[0] # Conexión ODBC Nacional
        self.rej_query = None # Query rechazos que se guarda en dataframe
        self.rej_folder = None # Carpeta donde se guardará el consolidado
        self.consolidation_file = None # Ruta del archivo consolidado con el query de rechazos
        self.df_canje = None # Dataframe que guardará la información de la hoja de TRN canje
        self.area_trns, self.df_areas = self._get_area_trns() # Diccionario para guardar las tablas de las hojas de archivo de transacciones áreas
    
    def _get_area_trns(self):
        """Obtiene parámetros de transacciones por área"""
        # Cargar hojas del archivo
        sheets = self.config['transacciones_areas']

        # Crear diccionario vacío para guardar dataframes
        area_trns = {}

        # Crear dataframe vacío para guardar la información de las áreas
        df_areas = pd.DataFrame()

        # Recorrer cada hoja y guardar dataframe en diccionario
        for sheet in sheets:
            # Cargar dataframe
            df = pd.read_excel(os.path.join(self.static, 'Transacciones áreas.xlsx'), sheet_name = sheet, dtype = str)

            # Agregar dataframe cosolidado áreas
            df_area = df.copy()
            df_area['Concatenar'] = df['Código aplicación'] + df['Código TRN']

            try:
                df_area['CTA CONTABLE'] = df['CUENTA CONTABLE']
            except:
                df_area['CTA CONTABLE'] = pd.Series(dtype = str)

            df_area = df_area[['Concatenar', 'CTA CONTABLE']]
            df_areas = pd.concat([df_areas, df_area], axis = 0).reset_index(drop = True)

            # Crear columna concatenado
            df[sheet] = df['Código aplicación'] + df['Código TRN'].astype(int).astype(str)

            # Quitar columnas de código de aplicación y transacción
            df = df[[col for col in df.columns if col not in ['Código aplicación', 'Código TRN']]]

            # Colocar sufijo a columnas
            for col in df.columns:
                if col != sheet:
                    df.rename(columns = {col: col + '_' + sheet}, inplace = True)

            # Agregar dataframe a diccionario
            area_trns[sheet] = df
        
        df_areas.to_excel('./out/ctas_cont.xlsx', index = False)

        return area_trns, df_areas

    def run_rej_query(self):
        """Ejecuta el query de rechazos"""
        # Crear carpeta para guardar consolidado en caso de que no exista
        self.rej_folder = os.path.join(self.paths['carpeta_rechazos'], self.end_date.strftime('%Y'), str(self.end_date.month) + '. ' + self.end_date.strftime('%B').upper(), self.settings['consolidado_rechazos']['nombre_carpeta'])

        if not os.path.exists(self.rej_folder):
            os.makedirs(self.rej_folder)
        
        # Crear dataframe vacío para guardar los querys
        df = pd.DataFrame()

        # Cargar query
        with io.open(os.path.join(self.sql_path, 'query_rechazos.sql'), mode = 'r', encoding = 'utf8') as f:
            sql = f.read()

        # Ejecutar query para cada una de las fechas
        current_date = self.start_date

        while current_date <= self.end_date:
            # Cambiar formato fecha
            date = current_date.strftime('%m%d')

            self.logger.info('Ejecutando query período ' + date)

            # Ejecutar query reemplazando fecha y guardar en dataframe
            df = pd.concat([df, self.query_AS400(sql.replace('MMDD', date), self.conn, True)[0]], axis = 0)

            # Sumar día al ciclo
            current_date += timedelta(days = 1)
        
        df = df.reset_index(drop = True)

        # Crear columna llave
        df['llave'] = df['Número identificación'].str.replace('^0+', '', regex = True) + 'r' + pd.Series(df.index.astype(str))

        self.rej_query = df
            
        self.rej_query.to_excel('./out/query_rechazos.xlsx', index = False)

        return None

    def transform_query(self):
        """Transforma, clasifica y aplica reglas de negocio a la información obtenida del query"""
        self.logger.info('Realizando merge con archivo areas transacciones')

        # Crear campo concatenar
        self.rej_query['Concatenar'] = self.rej_query['Código aplicación'] + self.rej_query['Código TRN'].astype(int).astype(str)

        # Hacer merge con cada una de las áreas
        for area, df in self.area_trns.items():
            try:
                self.rej_query = self.rej_query.merge(df, how = 'left', left_on = 'Concatenar', right_on = area, suffixes = (False, '_' + area), validate = 'm:1').reset_index(drop = True)
            except MergeError:
                msg = 'La hoja ' + area + ' en el archivo "Transacciones áreas" contiene duplicados'
                raise Exception(msg)
        
        # Hacer merge para obtener la cuenta contable
        self.rej_query = self.rej_query.merge(self.df_areas, how = 'left', on = 'Concatenar', validate = 'm:1')
            
        # Cambiar nombre columna cuenta contable medios de pago
        self.rej_query.rename(columns = {'CUENTA_MEDIOS DE PAGO': 'CTA CONTABLE Medios de pago'}, inplace = True)

        # Clasificar canje
        idxs = self.rej_query[self.rej_query['CANJE'].notnull()].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'CANJE'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'NO APLICADO'

        # Clasificar transacciones a validar
        idxs = self.rej_query[(self.rej_query['VALIDAR'].notnull()) & (self.rej_query['Instrucción_VALIDAR'] != 'TRN SUBSIDIO')].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'TRN PARA CONTABILIZAR'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'

        # Clasificar transacciones subsidio débito y crédito
        idxs = self.rej_query[(self.rej_query['VALIDAR'].notnull()) & (self.rej_query['Instrucción_VALIDAR'] == 'TRN SUBSIDIO') & (self.rej_query['Código DB o CR'] == 'C')].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'TRN SUBSIDIO'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'PENDIENTE'

        idxs = self.rej_query[(self.rej_query['VALIDAR'].notnull()) & (self.rej_query['Instrucción_VALIDAR'] == 'TRN SUBSIDIO') & (self.rej_query['Código DB o CR'] == 'D')].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'CXC'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'

        # Clasificar rechazos tipo 1
        idxs = self.rej_query[(self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código devolución'] == '1') & (self.rej_query['Número identificación'].isnull())].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'TIPO 1'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'PENDIENTE'

        # Clasificar rechazos en estado E
        idxs = self.rej_query[(self.rej_query['Clasificación'].isnull()) & (self.rej_query['Estado cuenta TR'].str.strip().isin(['E', 'P']))].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'EMBARGOS'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'PENDIENTE'

        # Definir áreas ################### validar si faltan áreas
        areas = ['PAGOS', 'MEDIOS DE PAGO', 'RECAUDOS', 'FINANCIACION']

        for area in areas:
            # Actualizar campo clasificación y estado rechazo cuentas embargadas
            idxs = self.rej_query[(self.rej_query[area].notnull()) & (self.rej_query['Clasificación'] == 'EMBARGOS')].index
            self.rej_query.loc[idxs, 'Clasificación'] = area
            self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'
        
        # Clasificar trxs que empiecen con iva comision o comision
        idxs = self.rej_query[(self.rej_query['Clasificación'] == 'EMBARGOS') & ((self.rej_query['Descripción TRN'].str.startswith('IVA COMIS')) | (self.rej_query['Descripción TRN'].str.startswith('COMIS')) | (self.rej_query['Descripción TRN'].str.startswith('IVA COBRO COMIS')))].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'CXC'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'

        # Clasificar rechazos en estado X
        idxs = self.rej_query[(self.rej_query['Clasificación'].isnull()) & ((self.rej_query['Estado cuenta TR'].str.strip().str.upper() == 'X') | (self.rej_query['Descripción TRN'].str.contains('AFC')))].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'RECHAZO CUENTA AFC'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'NO APLICADO'

        # Clasificar rechazos con valor cero
        idxs = self.rej_query[self.rej_query['Monto'].astype(float) == 0.0].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'NO APLICAR'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'NO APLICADO'

        # Cargar códigos devolución y hacer merge con consolidado
        dev_cod = self.settings['informe']['codigos_devolucion']
        df_dev = pd.DataFrame([[k, v] for k, v in dev_cod.items()], columns = ['cod_dev', 'desc_dev'])
        self.rej_query = pd.merge(self.rej_query, df_dev, how = 'left', left_on = 'Código devolución', right_on = 'cod_dev', validate = 'm:1')

        # Diligenciar explicación
        self.rej_query['Explicación'] = 'Rechazo fecha ' + self.rej_query['Fecha efectiva'].astype(int).astype(str) + ' transacción ' + self.rej_query['Código TRN'].astype(int).astype(str) + ' ' + self.rej_query['Descripción TRN'].str.strip() + ' cuenta ' + self.rej_query['Código aplicación'].map({'S': 'ahorros', 'D': 'corriente'}) + ' ' + self.rej_query['Número cuenta'].astype('int64').astype(str) + ' generado por la causal ' + self.rej_query['Código devolución'] + ' ' + self.rej_query['desc_dev']

        # Clasificar rechazos "VALIDAR", "TIPO 1" Y otras áreas
        for area in areas:
            # Condiciones de cuentas maestras
            conditions_1 = ((self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código aplicación'] == 'S') & (self.rej_query['Código DB o CR'] == 'C') & (self.rej_query['Estado cuenta TR'] == 'H'))
            conditions_2 = ((self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código aplicación'] == 'S') & (self.rej_query['Código DB o CR'] == 'D') & (self.rej_query['Estado cuenta TR'].isin(['H', 'D'])))
            conditions_3 = ((self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código aplicación'] == 'D') & (self.rej_query['Código DB o CR'] == 'D') & (self.rej_query['Estado cuenta TR'].isin(['G', 'D'])))

            # Actualizar campo clasificación y estado rechazo crédito de cuentas de ahorro
            idxs = self.rej_query[(conditions_1) & ((self.rej_query[area].notnull()))].index
            self.rej_query.loc[idxs, 'Clasificación'] = area
            self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'

            # Actualizar campo clasificación y estado rechazo débito de cuentas de ahorro
            idxs = self.rej_query[(conditions_2) & ((self.rej_query[area].notnull()))].index
            self.rej_query.loc[idxs, 'Clasificación'] = area
            self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'

            # Actualizar campo clasificación y estado rechazo débito de cuentas corriente
            idxs = self.rej_query[(conditions_3) & ((self.rej_query[area].notnull()))].index
            self.rej_query.loc[idxs, 'Clasificación'] = area
            self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'
        
        # Marcar el resto de cuentas maestras
        conditions_1 = ((self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código aplicación'] == 'S') & (self.rej_query['Código DB o CR'] == 'C') & (self.rej_query['Estado cuenta TR'] == 'H'))
        conditions_2 = ((self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código aplicación'] == 'S') & (self.rej_query['Código DB o CR'] == 'D') & (self.rej_query['Estado cuenta TR'].isin(['H', 'D'])))
        conditions_3 = ((self.rej_query['Clasificación'].isnull()) & (self.rej_query['Código aplicación'] == 'D') & (self.rej_query['Código DB o CR'] == 'D') & (self.rej_query['Estado cuenta TR'].isin(['G', 'D'])))

        idxs = self.rej_query[conditions_1].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'CXP'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'

        idxs = self.rej_query[(conditions_2) | (conditions_3)].index
        self.rej_query.loc[idxs, 'Clasificación'] = 'CXC'
        self.rej_query.loc[idxs, 'Estado rechazo'] = 'CONTABILIZAR'
            
        self.rej_query.to_excel('./out/merge.xlsx', index = False)

        # Crear dataframe hoja TRN canje
        self.df_canje = self.rej_query[(self.rej_query['Código aplicación'] == 'D') & (self.rej_query['Código TRN'].astype(int).isin([1267, 1268]))]

        # Quitar información de canje de la hoja de rechazos
        self.rej_query = self.rej_query[~self.rej_query.index.isin(self.df_canje.index)].reset_index(drop = True)

        return None

    def gen_consolidation_file(self):
        """Crea archivo de consolidación de rechazos"""
        # Configuración archivo rechazos
        orig_file = os.path.join(self.static, self.settings['consolidado_rechazos']['plantilla'])
        self.consolidation_file = os.path.join(self.rej_folder, self.settings['consolidado_rechazos']['nombre'].replace('AAAAMMDD', self.end_date.strftime('%Y%m%d')))
        rej_sheet_name = self.settings['consolidado_rechazos']['hoja_rechazos']
        canje_sheet_name = self.settings['consolidado_rechazos']['hoja_canje']

        # Crear una copia de la estructura
        shutil.copy(orig_file, self.consolidation_file)

        # Abrir archivo
        wb = xl.load_workbook(self.consolidation_file)

        # Nombre hojas
        ws_rej = wb[rej_sheet_name] # Hoja rechazos
        ws_canje = wb[canje_sheet_name] # Hoja canje

        # Llenar hojas
        self.fill_excel_sheet(ws_rej, self.rej_query)
        self.fill_excel_sheet(ws_canje, self.df_canje.reset_index(drop = True))

        # Guardar
        wb.save(self.consolidation_file)

        # Actualizar fecha procesamiento
        self.update_json(fecha_final_rechazos = datetime.strftime(self.end_date, '%Y%m%d'))

        return None

    def run(self):
        """Orquesta la ejecución de la clase"""
        self.run_rej_query()
        self.transform_query()
        self.gen_consolidation_file()

        return None
    