import requests

CLASIFICADOR_API_KEY = "gYERJxJHGv99eNvrs6Vdf63FJo77ebcX1RTdqBFe"
CLASIFICADOR_URL = "https://kbqui4zh4i.execute-api.us-east-1.amazonaws.com/V1/analizar-reclamo"

def clasificar(datos_caso: dict) -> dict:
    headers = {
        "x-api-key": CLASIFICADOR_API_KEY,
        "Content-Type": "application/json",
    }

    # La API espera el texto del reclamo como string plano
    response = requests.post(
        CLASIFICADOR_URL,
        headers=headers,
        json={"reclamo": datos_caso["texto"]},
        timeout=60,
    )
    response.raise_for_status()

    resultado = response.json()
    confianza_pct = round(resultado.get("factor_de_confianza", 0) * 100)
    print("=== ANÁLISIS DEL RECLAMO ===")
    print(f"Clasificación:  {resultado.get('clasificacion')}")
    print(f"Confianza: {confianza_pct}%")
    print(f"Justificación:  {resultado.get('justificacion')}")

    return resultado