import requests
from requests.auth import HTTPBasicAuth

# ── CREDENCIALES ────────────────────────────────────────────────────────────────
SEC_USERNAME = "ProyectoUC"
SEC_PASSWORD = "erMT0eA9QMVRCCle0Im1"

# ── CONFIGURACIÓN ───────────────────────────────────────────────────────────────
SEC_URL = "https://sec.custhelp.com/services/rest/connect/v1.3/analyticsReportResults"
ID_PENDIENTES = 105111
ID_REFERENCIA = "260128-000718"


# ── FUNCIÓN PRINCIPAL ────────────────────────────────────────────────────────────

def extraer(ID_REFERENCIA: int = ID_REFERENCIA, id_pendientes: int = ID_PENDIENTES) -> dict:
    # 1. Llamada a la API — solo se envía el ID de pendientes
    response = requests.post(
        SEC_URL,
        auth=HTTPBasicAuth(SEC_USERNAME, SEC_PASSWORD),
        json={"id": id_pendientes},
        timeout=30,
    )
    response.raise_for_status()

    data = response.json()
    rows = data.get("rows", [])
    columns = data.get("columnNames", [])

    # 2. Mapeo de columnas por nombre
    idx = {
        "id":                columns.index("ID de incidente"),
        "n_referencia":      columns.index("Nº referencia"),
        "fecha":             columns.index("Fecha creada"),
        "estado":            columns.index("Estado"),
        "cat3":              columns.index("Nivel de categoría 3"),
        "cat4":              columns.index("Nivel de categoría 4"),
        "empresa":           columns.index("Empresa"),
        "peticion":          columns.index("Petición Concreta"),
        "texto":             columns.index("Texto"),
        "respuesta_empresa": columns.index("Repuesta de Empresa (1era)"),
    }

    # 3. Buscar la fila que corresponde al ID_REFERENCIA
    fila = next(
        (row for row in rows if str(row[idx["n_referencia"]]) == str(ID_REFERENCIA)),
        None,
    )

    if fila is None:
        raise ValueError(
            f"Caso {ID_REFERENCIA} no encontrado en el reporte {id_pendientes} "
            f"({data.get('count', '?')} filas revisadas)."
        )

    # 4. Resultado estructurado
    resultado = {
        "id":                fila[idx["id"]],
        "n_referencia":      fila[idx["n_referencia"]],
        "fecha":             fila[idx["fecha"]],
        "estado":            fila[idx["estado"]],
        "cat3":              fila[idx["cat3"]],
        "cat4":              fila[idx["cat4"]],
        "empresa":           fila[idx["empresa"]],
        "peticion":          fila[idx["peticion"]],
        "texto":             fila[idx["texto"]],
        "respuesta_empresa": fila[idx["respuesta_empresa"]],
    }

    # 5. Log de consola
    print("=== CASO EXTRAÍDO ===")
    print(f"ID: {resultado['id']} | Ref: {resultado['n_referencia']} | Fecha: {resultado['fecha']}")
    print(f"Empresa: {resultado['empresa']} | Estado: {resultado['estado']}")
    print(f"Categoría 3: {resultado['cat3']}")
    print(f"Categoría 4: {resultado['cat4']}")
    print(f"Petición: {resultado['texto']}")

    return resultado

if __name__ == "__main__":
    resultado = extraer(ID_REFERENCIA=ID_REFERENCIA, id_pendientes=ID_PENDIENTES)
    print(resultado)