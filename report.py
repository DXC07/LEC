from functions import Functions
import os
import pandas as pd
import shutil
from datetime import datetime
import numpy as np
import win32com.client
import io

class Report(Functions):

    def __init__(self, date):
        super().__init__()
        self.date = date # Fecha del último día de los rechazos gestionados
        self.df_areas = self._get_area_info() # Dataframe con la información de las áreas
        self.df_cons = self.read_consolidation_file() # Consolidado rechazos
        self.mail_file = os.path.join(self.paths['onedrive'], self.settings['archivo_mails_enviados'])
    
    def _get_area_info(self):
        """Obtiene DataFrame con la información de las áreas parametrizadas"""
        # Definir dataframe vacío
        df = pd.DataFrame()

        # Áreas a tener en cuenta
        areas = self.config["areas_informe"]

        # Recorrer cada área
        for area in areas:
            df_area = pd.read_excel(os.path.join(self.static, 'Transacciones áreas.xlsx'), sheet_name = area, dtype = str)
            df_area['Concatenar'] = df_area['Código aplicación'] + df_area['Código TRN']
            df_area['area'] = area
            df_area = df_area[['Concatenar', 'area', 'SECCIÓN RESPONSABLE']]
            df = pd.concat([df, df_area], axis = 0).reset_index(drop = True)

        # Quitar duplicados
        df.drop_duplicates(inplace = True)

        return df
    
    def read_consolidation_file(self):
        """Carga la información que se encuentra en el consolidado del día"""
        # Definir ruta archivo rechazos
        self.rej_folder = os.path.join(self.paths['carpeta_rechazos'], self.date.strftime('%Y'), str(self.date.month) + '. ' + self.date.strftime('%B').upper(), self.settings['consolidado_rechazos']['nombre_carpeta'])
        self.consolidation_file = os.path.join(self.rej_folder, self.settings['consolidado_rechazos']['nombre'].replace('AAAAMMDD', self.date.strftime('%Y%m%d')))

        # Cargar hoja de rechazos en Dataframe
        df = pd.read_excel(self.consolidation_file, sheet_name = self.settings['consolidado_rechazos']['hoja_rechazos'])

        # Cargar códigos devolución y hacer merge con consolidado
        dev_cod = self.settings['informe']['codigos_devolucion']
        df_dev = pd.DataFrame([[k, v] for k, v in dev_cod.items()], columns = ['cod_dev', 'desc_dev'])
        df = pd.merge(df, df_dev, how = 'left', left_on = 'Código devolución', right_on = 'cod_dev', validate = 'm:1')

        # Hacer merge con archivo de áreas
        df = pd.merge(df, self.df_areas, how = 'left', on = 'Concatenar', validate = 'm:1')

        # # Quitar canje
        # df = df[df['Clasificación'] != 'CANJE'].reset_index(drop = True)

        return df
    
    def gen_report(self):
        """Genera el reporte que se publica"""
        #Intentar borrar archivo de mails enviados
        try:
            os.remove(self.mail_file)
        except FileNotFoundError:
            pass

        # Crear dataframe para el informe
        df_report = pd.DataFrame()

        # Definir cada columna
        df_report['Código de aplicación'] = self.df_cons['Código aplicación']
        df_report['Sucursal libros'] = self.df_cons['Sucursal libros']
        df_report['Sucursal ingreso'] = self.df_cons['Sucursal ingreso']
        df_report['Fecha efectiva'] = self.df_cons['Fecha efectiva']
        df_report['Código TRN'] = self.df_cons['Código TRN']
        df_report['Descripción TRN'] = self.df_cons['Descripción TRN']
        df_report['Código DB o CR'] = self.df_cons['Código DB o CR']
        df_report['Monto'] = self.df_cons['Monto']
        df_report['Código devolución'] = self.df_cons['Código devolución'] + ' ' + self.df_cons['desc_dev']
        df_report['Número cuenta'] = self.df_cons['Número cuenta']
        df_report['Fecha vinculación'] = self.df_cons['Fecha vinculación']
        df_report['Número cheque o serie'] = self.df_cons['Número cheque o serie']
        df_report['Estado cuenta'] = self.df_cons['Estado cuenta']
        df_report['Estado cuenta TR'] = self.df_cons['Estado cuenta TR']
        df_report['Código TRN rechazo'] = self.df_cons['Código TRN rechazo']
        df_report['Código rastreo'] = self.df_cons['Código rastreo']
        df_report['Tipo identificación'] = self.df_cons['Tipo identificación']
        df_report['Número identificación'] = self.df_cons['Número identificación']
        df_report['Nombre'] = self.df_cons['Nombre']
        df_report['Segmento'] = self.df_cons['Segmento']
        df_report['Tipo cliente'] = self.df_cons['Tipo cliente']
        df_report['Explicación'] = self.df_cons['Explicación']
        df_report['Fecha proceso de rechazos'] = datetime.now().date()
        
        # Definir áreas
        areas = self.config["areas_informe"]
        
        # Crear función para la columna resultado proceso
        def func(row):
            if row['Clasificación'] in areas:
                return 'CONTABILIZADO A ' + row['Clasificación']
            elif row['Clasificación'] in ['CXC', 'CXP']:
                return 'CONTABILIZADO EN ' + row['Clasificación']
            elif row['Clasificación'] == 'NO APLICAR':
                return 'NO APLICADO'
            elif row['Clasificación'] in ['LECTURA TIEMPO REAL', 'LECTURA BATCH', 'EMBARGOS LECTURA TIEMPO REAL', 'TRN SUBSIDIO LECTURA TIEMPO REAL', 'F', 'R']:
                return 'APLICADO EN LA CUENTA DEL CLIENTE'
            elif row['Clasificación'] == 'CANJE':
                return 'RECHAZO DE CANJE, NO APLICADO'
            elif row['Clasificación'] == 'RECHAZO CUENTA AFC':
                return 'RECHAZO CUENTA AFC, NO APLICADO'
            else:
                return row['Clasificación']
        
        df_report['Resultado del proceso'] = self.df_cons.apply(func, axis = 1)
        df_report['Gerencia dueña de transacción'] = self.df_cons.apply(lambda row: row['Clasificación'] if row['Clasificación'] in areas else np.nan, axis = 1)   
        df_report['Sección responsable'] = self.df_cons.apply(lambda row: row['SECCIÓN RESPONSABLE'] if row['Clasificación'] in areas else np.nan, axis = 1)

        # Crear informe konecta con cuenta cifrada
        df_konecta = df_report.copy()
        df_konecta['cta_cifrada'] = df_konecta['Número cuenta'].astype('int64').astype(str).str.slice(-4).astype(int)
        df_konecta['Explicación'] = df_konecta.apply(lambda row: row['Explicación'].replace(str(int(row['Número cuenta'])), str(row['cta_cifrada'])), axis = 1)
        df_konecta['Número cuenta'] = df_konecta['cta_cifrada']
        df_konecta.drop(columns = ['cta_cifrada'], inplace = True)

        df_report.to_excel('./out/informe.xlsx', index = False)
        df_konecta.to_excel('./out/informe_konecta.xlsx', index = False)

        # Definir ruta informes
        report_file = os.path.join(self.paths['onedrive'], 'Informes', self.settings['informe']['nombre'].replace('AAAAMM', datetime.now().strftime('%Y%m')))
        konecta_report_file = os.path.join(self.paths['onedrive'], 'Informes', self.settings['informe']['nombre_konecta'].replace('AAAAMM', datetime.now().strftime('%Y%m')))

        fill = True

        # Validar si ya existe informe normal
        try:
            df_r = pd.read_excel(report_file)

            # Validar si ya existen registros con la fecha de hoy
            max_date = df_r['Fecha proceso de rechazos'].max().date()

            if max_date < datetime.now().date():
                row = df_r.shape[0] + 2
            else:
                fill = False
        except FileNotFoundError:
            shutil.copy(os.path.join(self.static, 'Estructura informe rechazos depósitos.xlsx'), report_file)
            row = 2
        
        # Llenar informe
        if fill:
            self.update_excel_file(report_file, 'Hoja1', df_report, False, row)
            self.logger.info('Se generó informe rechazos depósitos')
        else:
            self.logger.warning('Ya existen registros del día de hoy en el informe rechazos depósitos')

        fill = True

        # Validar si ya existe informe konecta
        try:
            df_r = pd.read_excel(konecta_report_file)

            # Validar si ya existen registros con la fecha de hoy
            max_date = df_r['Fecha proceso de rechazos'].max().date()

            if max_date < datetime.now().date():
                row = df_r.shape[0] + 2
            else:
                fill = False
            row = df_r.shape[0] + 2
        except FileNotFoundError:
            shutil.copy(os.path.join(self.static, 'Estructura informe rechazos depósitos.xlsx'), konecta_report_file)
            row = 2
        
        # Llenar informe
        if fill:
            self.update_excel_file(konecta_report_file, 'Hoja1', df_konecta, False, row)
            self.logger.info('Se generó informe rechazos Konecta')
        else:
            self.logger.warning('Ya existen registros del día de hoy en el informe rechazos Konecta')

        return None

    def send_notifications(self):
        """Envía notificación de cargue de informe a los usuarios parametrizados"""
        self.logger.info('Esperando a que se actualice el informe en Sharepoint')
        
        # Esperar a que se actualice el informe
        while not os.path.exists(self.mail_file):
            pass

        self.logger.info('Enviando notificación de cargue de informe')

        # Crear sesión de outlook
        outlook = win32com.client.Dispatch('outlook.application')

        # Cargar plantilla
        with io.open(os.path.join(self.static, 'notificacion_areas.html'), mode = 'r', encoding = 'utf8') as f:
            hbody = f.read()
        
        # Cargar usuarios
        users = [user + "@bancolombia.com.co" for user in self.config["usuarios_notificacion_informe"]]
        to = ';'.join(users)
        self.logger.debug('Se envió notificación a ' + to)

        self.send_mail(outlook, to, 'Informe rechazos de depósitos', False, True, hbody = hbody)

        return None
    
    def run(self):
        """Orquesta la ejecución de la clase"""
        self.gen_report()
        self.send_notifications()
        
        return None