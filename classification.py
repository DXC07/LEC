from functions import Functions
import os
import pandas as pd
import io
import openpyxl as xl

class Classification(Functions):

    def __init__(self, credentials, date):
        super().__init__()
        self.nac_user = credentials['login_NACIONAL_user'] # Usuario nacional
        self.nac_pwd = credentials['login_NACIONAL_pwd'] # Contraseña nacional
        self.date = date # Fecha del último día de los rechazos gestionados
        self.conn = self.connect_odbc('AS400', maq = 'NACIONAL', user = self.nac_user, pwd = self.nac_pwd)[0] # Conexión ODBC Nacional
        self.df_areas = self._get_area_info() # Dataframe con la información de las áreas

    def _gen_state_sld_tr(self, accounts):
        """Retorna dataframe con cuentas y su estado y saldo en tiempo real"""
        # Convetir cuentas en una lista y en formato str
        accounts = accounts.to_list()
        accounts = [str(a) for a in accounts]

        # Leer consulta tr
        with io.open(os.path.join(self.sql_path, 'saldo_estado_TR.sql'), mode = 'r', encoding = 'utf8') as f:
            sql_tr = f.read()
        
        # Crear dataframe vacío para guardar la respuesta
        df = pd.DataFrame()

        # Consultar de a 100 cuentas
        for i in range(0, len(accounts) + 1, 100):
            filtered_accs = accounts[i: i + 100]
            sql_accs = ','.join(filtered_accs)

            sql = sql_tr.replace('cuentas', sql_accs)

            df = pd.concat([df, self.query_AS400(sql, self.conn)[0]], axis = 0)
        
        # Quitar duplicados
        df.drop_duplicates(inplace = True)

        df.to_excel('./out/saldo_tr.xlsx', index = False)

        return df
    
    def _get_area_info(self):
        """Obtiene DataFrame con la información de las áreas parametrizadas"""
        # Definir dataframe vacío
        df = pd.DataFrame()

        # Áreas a tener en cuenta
        areas = ['PAGOS', 'RECAUDOS', 'MEDIOS DE PAGO', 'FINANCIACION', 'OPERACION INMOBILIARIA', 'COMERCIO INTERNACIONAL', 'CONCILIACION SUFI']

        # Recorrer cada área
        for area in areas:
            df_area = pd.read_excel(os.path.join(self.static, 'Transacciones áreas.xlsx'), sheet_name = area, dtype = str)
            df_area['Concatenar'] = df_area['Código aplicación'] + df_area['Código TRN']
            df_area['area'] = area
            df_area = df_area[['Concatenar', 'area']]
            df = pd.concat([df, df_area], axis = 0).reset_index(drop = True)

        # Quitar duplicados
        df.drop_duplicates(inplace = True)

        return df

    def update_consolidation_file(self):
        """Carga la información que se encuentra en el consolidado del día"""
        # Definir ruta archivo
        self.rej_folder = os.path.join(self.paths['carpeta_rechazos'], self.date.strftime('%Y'), str(self.date.month) + '. ' + self.date.strftime('%B').upper(), self.settings['consolidado_rechazos']['nombre_carpeta'])
        self.consolidation_file = os.path.join(self.rej_folder, self.settings['consolidado_rechazos']['nombre'].replace('AAAAMMDD', self.date.strftime('%Y%m%d')))

        # Cargar hoja de rechazos en Dataframe
        df = pd.read_excel(self.consolidation_file, sheet_name = self.settings['consolidado_rechazos']['hoja_rechazos'])

        # Definir áreas ################### validar si faltan áreas
        areas = ['PAGOS', 'MEDIOS DE PAGO', 'RECAUDOS', 'CANJE', 'FINANCIACION']
        new_areas = ['OPERACION INMOBILIARIA', 'COMERCIO INTERNACIONAL', 'CONCILIACION SUFI']

        # Hacer merge para saber si las trx corresponden a nuevas áreas
        df = pd.merge(df, self.df_areas, how = 'left', on = 'Concatenar', validate = 'm:1')
        df.to_excel('./out/trx_nuevas.xlsx', index= False)

        # Clasificar nuevas áreas
        idxs = df[((df['Clasificación'].isnull()) | (df['Estado rechazo'] == 'PENDIENTE')) & (df['area'].isin(new_areas))  & (df['Código devolución'] == 'O')].index 
        df.loc[idxs, 'Clasificación'] = df.loc[idxs, 'area']
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        # Clasificar rechazos "VALIDAR", "TIPO 1" y "APLICAR MANUAL" a otras áreas
        for area in areas:
            # Actualizar campo clasificación y estado rechazo
            idxs = df[(df['Clasificación'].isin(['VALIDAR', 'TIPO 1', 'APLICAR MANUAL'])) & (df[area].notnull())].index
            df.loc[idxs, 'Clasificación'] = area
            df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"
        
        # Validar último ciclo transacciones pendientes ########################
        for area in areas:
            # Actualizar campo clasificación y estado rechazo
            idxs = df[((df['Clasificación'].isnull()) | (df['Estado rechazo'] == 'PENDIENTE')) & (df[area].notnull()) & (df['Código DB o CR'] == 'D')].index
            df.loc[idxs, 'Clasificación'] = area
            df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"
        
        # Calcular saldo y estado TR cuentas pendientes
        accounts = df[df['Clasificación'].isnull()]['Número cuenta']
        df_accounts = self._gen_state_sld_tr(accounts)

        # Cambiar formato columnas
        df['Número cuenta'] = df['Número cuenta'].astype('int64')
        df_accounts['SDCUENTA'] = df_accounts['SDCUENTA'].astype('int64')

        # Hacer merge con el consolidado
        df = df.merge(df_accounts, how = 'left', left_on = ['Código aplicación', 'Número cuenta'], right_on = ['SDTIPOCTA', 'SDCUENTA'], validate = 'm:1')

        # Actualizar saldo y estado
        idxs = df[df['SDCUENTA'].notnull()].index
        df.loc[idxs, 'Saldo disponible TR'] = df.loc[idxs, 'Saldo TR']
        df.loc[idxs, 'Estado cuenta TR'] = df.loc[idxs, 'SDESTADO']

        # Clasificar cuentas en estados cancelatorios o bloqueo y que el saldo no cubre el monto del rechazo
        canc_states = self.config['estados_cancelatorios_bloqueo']
        idxs = df[(df['Código DB o CR'] == 'D') & (df['Clasificación'].isnull()) & (df['Estado cuenta TR'].isin(canc_states))].index
        df.loc[idxs, 'Clasificación'] = "CXC"
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        # Clasificar cuentas de ahorro en estados vigentes y que el saldo no cubre el monto del rechazo
        act_states = self.config['estados_vigentes']
        idxs = df[(df['Código DB o CR'] == 'D') & (df['Clasificación'].isnull()) & (df['Código aplicación'].str.strip() == 'S') & (df['Estado cuenta TR'].isin(act_states))].index
        df.loc[idxs, 'Clasificación'] = "LECTURA BATCH"
        df.loc[idxs, 'Estado rechazo'] = "PENDIENTE"

        # Clasificar cuentas corriente que no tienen saldo, cupo de sobregiro o este se encuentra bloqueado
        idxs = df[(df['Código DB o CR'] == 'D') & (df['Clasificación'].isnull()) & (df['Código aplicación'].str.strip() == 'D')].index
        df.loc[idxs, 'Clasificación'] = "CXC"
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        # Clasificar rechazos crédito que no se puedan gestionar
        idxs = df[(df['Código DB o CR'] == 'C') & (df['Clasificación'].isnull())].index
        df.loc[idxs, 'Clasificación'] = "CXP"
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        # Clasificar rechazos embargos no autorizados
        idxs = df[df['Clasificación'] == 'EMBARGOS'].index
        df.loc[idxs, 'Clasificación'] = "CXC"
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        # Clasificar rechazos subsidios naturaleza débito
        idxs = df[(df['Clasificación'] == 'TRN SUBSIDIO') & (df['Código DB o CR'] == 'D')].index
        df.loc[idxs, 'Clasificación'] = "CXC"
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        # Clasificar rechazos subsidios naturaleza débito
        idxs = df[(df['Clasificación'] == 'TRN SUBSIDIO') & (df['Código DB o CR'] == 'C')].index
        df.loc[idxs, 'Clasificación'] = "CXP"
        df.loc[idxs, 'Estado rechazo'] = "CONTABILIZAR"

        self.df_rej = df

        # Actualizar archivo consolidado
        self.update_excel_file(self.consolidation_file, self.settings['consolidado_rechazos']['hoja_rechazos'], df)

        return None

    def run(self):
        """Orquesta la ejecución de la clase"""
        self.update_consolidation_file()
        
        return None