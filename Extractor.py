import requests
import pandas as pd

class Extractor_Reclamo:
    def __init__(self):
        self.url_api = "https://sec.custhelp.com/services/rest/connect/v1.3/analyticsReportResults"
        self.user = "ProyectoUC"
        self.password = "erMT0eA9QMVRCCle0Im1"

    def get_reclamo(self, id_reporte, paso=10000):
        self.id_reporte = id_reporte
        try:
            offset = 0
            datos_acumulados = []
            columnas = []

            while True:
                body = {
                    "id": id_reporte,
                    "offset": offset
                }
                headers = {
                    "Content-Type": "application/json",
                    "OSvC-CREST-Application-Context": "Run report"
                }
                response = requests.post(
                    self.url_api, 
                    auth=(self.user, self.password),
                    headers=headers,
                    json=body
                )
                if response.status_code != 200:
                    raise ValueError(f"Error en la solicitud: {response.status_code} - {response.text}")
                
                content = response.json()
                filas = content.get("rows", [])
                
                if not filas:
                    break
                if not columnas:
                    columnas = content.get("columnNames", [])
                
                datos_acumulados.extend(filas)
                offset += paso

            return pd.DataFrame(datos_acumulados, columns=columnas)
            
        except Exception as e:
            print(f"Error al procesar el reporte: {e}")
            return None

# Instanciación y uso:
extractor = Extractor_Reclamo()
df_reclamos = extractor.get_reclamo(id_reporte=105111)

df_reclamos.to_csv(f"info_reclamo_{extractor.id_reporte}.csv", index=False, encoding="utf-8-sig")
print("Archivo CSV extraído correctamente")