import pandas as pd
from sodapy import Socrata
from utils.CleanData import clean_url, clean_date

from .logging_setup import get_logger

logger = get_logger(__name__)

class OpenDataTools:
    def __init__(self):
        self.client = Socrata("www.datos.gov.co", None)
        
        # Identificadores de SECOP I y SECOP II
        self.secopi_contratos = "f789-7hwg"
        self.secopii_contratos = "jbjy-vk9h"
        self.secopii_procesos = "p6dx-8zbt"
        self.tvec_consolidado = "rgxm-mmea"

    def get_cedulas(cedulas, i, j):
        cadena_final = '('
        for doc in cedulas[i:j-1]:
            cadena_final += "'" + f"{doc}', "

        cadena_final += f"'{cedulas[j]}')"
        return cadena_final
    
    def clean_results_S1(self, results):
        try:
            # Limpieza de la URL de los contratos
            results['ruta_proceso_en_secop_i'] = results['ruta_proceso_en_secop_i'].astype(str)
            results['ruta_proceso_en_secop_i'] = results['ruta_proceso_en_secop_i'].apply(clean_url)
            # Ajuste sobre las fechas
            results['fecha_de_cargue_en_el_secop'] = results['fecha_de_cargue_en_el_secop'].astype(str)
            results['fecha_de_cargue_en_el_secop'] = results['fecha_de_cargue_en_el_secop'].apply(clean_date)
            results['fecha_de_firma_del_contrato'] = results['fecha_de_firma_del_contrato'].astype(str)
            results['fecha_de_firma_del_contrato'] = results['fecha_de_firma_del_contrato'].apply(clean_date)
            results['fecha_fin_ejec_contrato'] = results['fecha_fin_ejec_contrato'].astype(str)
            results['fecha_fin_ejec_contrato'] = results['fecha_fin_ejec_contrato'].apply(clean_date)
            print(f'Se identificaron {results.shape[0]} registros en SECOP I.')
        except:
            print('No se encontraron datos para el SECOP I.')
        
        return results
    
    def clean_results_S2(self, results):
        try:
            # Limpieza de la URL de los contratos
            results['urlproceso'] = results['urlproceso'].astype(str)
            results['urlproceso'] = results['urlproceso'].apply(clean_url)
            # Ajuste sobre las fechas
            results['fecha_de_firma'] = results['fecha_de_firma'].astype(str)
            results['fecha_de_firma'] = results['fecha_de_firma'].apply(clean_date)
            print(f'Se identificaron {results.shape[0]} registros en SECOP II.')
        except:
            print('No se encontraron datos para el SECOP II.')
        return results
    
    def clean_results_tvec(self, results):
        try:
            print(f'Se identificaron {results.shape[0]} registros en TVEC.')
        except:
            print('No se encontraron datos para TVEC.')
    
    def get_contratos_por_cedulas(self, cedulas, limit=10000):
        """
        Las cédulas deben estar en el formato '("30579584")'
        """
        query_s1 = f"""
            select 
                uid, numero_de_contrato, id_adjudicacion, anno_cargue_secop, 
                fecha_de_cargue_en_el_secop, anno_firma_contrato, fecha_de_firma_del_contrato,
                fecha_fin_ejec_contrato, tipo_de_contrato, modalidad_de_contratacion, 
                causal_de_otras_formas_de, estado_del_proceso, objeto_del_contrato_a_la, 
                detalle_del_objeto_a_contratar, tipo_identifi_del_contratista, 
                identificacion_del_contratista, nom_razon_social_contratista, 
                tipo_doc_representante_legal, identific_representante_legal, 
                nombre_del_represen_legal, nombre_entidad, nit_de_la_entidad, 
                departamento_entidad, municipio_entidad, valor_contrato_con_adiciones, 
                ruta_proceso_en_secop_i
            where
                (identificacion_del_contratista in {cedulas} or 
                identific_representante_legal in {cedulas})
            limit
            {limit}
            """
        logger.debug(f'Query on SI: {query_s1}')

        contratos_s1 = self.client.get(self.secopi_contratos, content_type="json", 
                                       query=query_s1)

        query_s2 = f"""
            select 
                id_contrato, fecha_de_firma, tipo_de_contrato, 
                modalidad_de_contratacion, estado_contrato, objeto_del_contrato, 
                tipodocproveedor, documento_proveedor, proveedor_adjudicado, 
                tipo_de_identificaci_n_representante_legal, identificaci_n_representante_legal, 
                nombre_representante_legal, nombre_entidad, nit_entidad, departamento, ciudad, 
                valor_del_contrato, urlproceso
            where
                (documento_proveedor in {cedulas} or 
                identificaci_n_representante_legal in {cedulas})
            limit
            {limit}
            """
        # Moved columns: fecha_de_fin_de_ejecucion
        contratos_s2 = self.client.get(self.secopii_contratos, content_type="json", 
                                       query=query_s2)
        
        # query_tvec = f"""
        #     select 
        #         identificador_de_la_orden, a_o, fecha, agregacion, estado, items, ciudad, 
        #         total, entidad, nit_entidad, solicitante, proveedor
        #     where
        #         nit_proveedor in {cedulas}
        #     limit
        #     {limit}
        #     """
        # contratos_tvec = self.client.get(self.tvec_consolidado, content_type="json", 
        #                                  query=query_tvec)
        
        # Limpieza de los contratos
        contratos_s1 = self.clean_results_S1(pd.DataFrame.from_dict(contratos_s1))
        contratos_s2 = self.clean_results_S2(pd.DataFrame.from_dict(contratos_s2))
        # contratos_tvec = self.clean_results_tvec(pd.DataFrame.from_dict(contratos_tvec))

        return {'SECOP_I': contratos_s1, 'SECOP_II': contratos_s2} #, 'TVEC': contratos_tvec}
    
    def get_contratos_por_nits(self, nits, limit=1000):
        """
        Los nits deben estar en el formato '("30579584")'
        """
        query_s1 = f"""
            select 
                uid, numero_de_contrato, id_adjudicacion, anno_cargue_secop, 
                fecha_de_cargue_en_el_secop, anno_firma_contrato, fecha_de_firma_del_contrato, 
                fecha_fin_ejec_contrato, tipo_de_contrato, modalidad_de_contratacion, 
                causal_de_otras_formas_de, estado_del_proceso, objeto_del_contrato_a_la, 
                detalle_del_objeto_a_contratar, tipo_identifi_del_contratista, 
                identificacion_del_contratista, nom_razon_social_contratista, 
                tipo_doc_representante_legal, identific_representante_legal, 
                nombre_del_represen_legal, nombre_entidad, nit_de_la_entidad, 
                departamento_entidad, municipio_entidad, valor_contrato_con_adiciones, 
                ruta_proceso_en_secop_i
            where
                nit_de_la_entidad in {nits}
            limit
            {limit}
            """
        contratos_s1 = self.client.get(self.secopi_contratos, content_type="json", 
                                       query=query_s1)

        query_s2 = f"""
            select 
                id_contrato, fecha_de_firma, tipo_de_contrato, 
                modalidad_de_contratacion, estado_contrato, objeto_del_contrato, 
                tipodocproveedor, documento_proveedor, proveedor_adjudicado, 
                tipo_de_identificaci_n_representante_legal, identificaci_n_representante_legal, 
                nombre_representante_legal, nombre_entidad, nit_entidad, departamento, ciudad, 
                valor_del_contrato, urlproceso
            where
                nit_entidad in {nits}
            limit
            {limit}
            """
        # Moved columns: fecha_de_fin_de_ejecucion
        contratos_s2 = self.client.get(self.secopii_contratos, content_type="json", 
                                       query=query_s2)
        
        # query_tvec = f"""
        #     select 
        #         identificador_de_la_orden, a_o, fecha, agregacion, estado, items, ciudad, 
        #         total, entidad, nit_entidad, solicitante, proveedor, nit_proveedor
        #     where
        #         nit_entidad in {nits}
        #     limit
        #     {limit}
        #     """
        # contratos_tvec = self.client.get(self.tvec_consolidado, content_type="json", 
        #                                  query=query_tvec)
        
        # Limpieza de los contratos
        contratos_s1 = self.clean_results_S1(pd.DataFrame.from_dict(contratos_s1))
        contratos_s2 = self.clean_results_S2(pd.DataFrame.from_dict(contratos_s2))
        # contratos_tvec = self.clean_results_tvec(pd.DataFrame.from_dict(contratos_tvec))

        return {'SECOP_I': contratos_s1, 'SECOP_II': contratos_s2} #, 'TVEC': contratos_tvec}
    
    def get_contratos_por_nombre_proveedor(self, nombres, limit=1000):
        """
        Los nits deben estar en el formato '("nombre")'
        """
        query_s1 = f"""
            select 
                uid, numero_de_contrato, id_adjudicacion, anno_cargue_secop, 
                fecha_de_cargue_en_el_secop, anno_firma_contrato, fecha_de_firma_del_contrato, 
                fecha_fin_ejec_contrato, tipo_de_contrato, modalidad_de_contratacion, 
                causal_de_otras_formas_de, estado_del_proceso, objeto_del_contrato_a_la, 
                detalle_del_objeto_a_contratar, tipo_identifi_del_contratista, 
                identificacion_del_contratista, nom_razon_social_contratista, 
                tipo_doc_representante_legal, identific_representante_legal, 
                nombre_del_represen_legal, nombre_entidad, nit_de_la_entidad, 
                departamento_entidad, municipio_entidad, valor_contrato_con_adiciones, 
                ruta_proceso_en_secop_i
            where
                nom_razon_social_contratista in {nombres} or nombre_del_represen_legal in {nombres}
            limit
            {limit}
            """
        contratos_s1 = self.client.get(self.secopi_contratos, content_type="json", 
                                       query=query_s1)

        query_s2 = f"""
            select 
                id_contrato, fecha_de_firma, tipo_de_contrato, 
                modalidad_de_contratacion, estado_contrato, objeto_del_contrato, 
                tipodocproveedor, documento_proveedor, proveedor_adjudicado, 
                tipo_de_identificaci_n_representante_legal, identificaci_n_representante_legal, 
                nombre_representante_legal, nombre_entidad, nit_entidad, departamento, ciudad, 
                valor_del_contrato, urlproceso
            where
                proveedor_adjudicado in {nombres} or nombre_representante_legal in {nombres}
            limit
            {limit}
            """
        # Moved columns: fecha_de_fin_de_ejecucion
        contratos_s2 = self.client.get(self.secopii_contratos, content_type="json", 
                                       query=query_s2)
        
        query_tvec = f"""
            select 
                identificador_de_la_orden, a_o, fecha, agregacion, estado, items, ciudad, 
                total, entidad, nit_entidad, solicitante, proveedor
            where
                proveedor in {nombres}
            limit
            {limit}
            """
        contratos_tvec = self.client.get(self.tvec_consolidado, content_type="json", 
                                         query=query_tvec)
        
        # Limpieza de los contratos
        contratos_s1 = self.clean_results_S1(pd.DataFrame.from_dict(contratos_s1))
        contratos_s2 = self.clean_results_S2(pd.DataFrame.from_dict(contratos_s2))
        contratos_tvec = self.clean_results_tvec(pd.DataFrame.from_dict(contratos_tvec))

        return {'SECOP_I': contratos_s1, 'SECOP_II': contratos_s2, 'TVEC': contratos_tvec}