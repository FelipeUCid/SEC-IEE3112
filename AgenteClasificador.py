import re
import ollama

class AgenteClasificadorSEC:
    def __init__(self, raw_text):
        self.raw_text = raw_text
        self.prompt_inicial = (
            "Eres un software clasificador de reclamos para una empresa. "
            "Analiza el reclamo del usuario y clasifícalo SOLO en una de estas opciones: "
            "[FACTURACION, LOGISTICA, SOPORTE_TECNICO, DEVOLUCIONES]. "
            "Responde únicamente con la palabra de la categoría, sin textos adicionales, sin puntos y sin saludos."
        )

        # Palabras irrelevantes
        self._frases_irrelevantes = [
            r"\b(hola|buenos días|buenas tardes|buenas noches|estimados?|saludos?)\b",
            r"\b(por favor|muchas gracias|gracias|de antemano|quedo atento|quedo a la espera)\b",
            r"\b(atte|atentamente|cordialmente|me despido)\b",
            r"\b(espero (su|una) (pronta )?respuesta)\b",
        ]

        # RUT chileno
        self._patron_rut = r"\b\d{1,2}\.?\d{3}\.?\d{3}-[\dkK]\b"

        self._patron_nombre_explicito = (
            r"(?i)"
            r"(?:me llamo|soy|mi nombre es)\s+"
            r"([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+"
            r"(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){0,3})"
        )

        # Correos electrónicos
        self._patron_email = r"\b[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}\b"

        # Teléfonos chilenos: +56 9 1234 5678, 9-1234-5678, (2) 2345-6789, etc.
        self._patron_telefono = (
            r"(\+?56\s?)?"           # Prefijo país opcional
            r"(\(?\d{1,2}\)?\s?)?"  # Código de área opcional
            r"\d{4}[\s\-]?\d{4}"    # 8 dígitos con separador opcional
        )

    def procesar_texto(self) -> str:
        """
        Limpia el texto del reclamo:
          1. Elimina frases de saludo/cierre irrelevantes para la clasificación.
          2. Censura RUTs, nombres explícitos, emails y teléfonos.
          3. Normaliza espacios y saltos de línea.
        """

        # 1. Eliminar frases irrelevantes (case-insensitive)
        for patron in self._frases_irrelevantes:
            texto = re.sub(patron, "", self.raw_text, flags=re.IGNORECASE)

        # 2. Censurar datos sensibles

        # Nombres explícitos (primero, antes de quitar el contexto)
        texto = re.sub(
            self._patron_nombre_explicito,
            lambda m: m.group(0).replace(m.group(1), "[NOMBRE CENSURADO]"),
            texto
        )

        # RUT chileno
        texto = re.sub(self._patron_rut, "[RUT CENSURADO]", texto)

        # Emails
        texto = re.sub(self._patron_email, "[EMAIL CENSURADO]", texto)

        # Teléfonos
        texto = re.sub(self._patron_telefono, "[TELÉFONO CENSURADO]", texto)

        # 3. Normalizar espacios: colapsar múltiples espacios/saltos en uno solo
        texto = re.sub(r"\s+", " ", texto).strip()

        return texto

    def clasificar_prompt(self):
        reclamo_procesado = self.procesar_texto()
        response = ollama.chat(
            model="llama3.2",
            messages=[
                {"role": "system", "content": self.prompt_inicial},
                {"role": "user", "content": f"Reclamo: {reclamo_procesado}"}
            ],
            options={
                "temperature": 0.0
            }
        )
        return response["message"]["content"].strip()

    
    