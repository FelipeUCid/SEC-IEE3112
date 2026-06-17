import { BedrockRuntimeClient, InvokeModelCommand } from "@aws-sdk/client-bedrock-runtime";
import crypto from "node:crypto";

const bedrock = new BedrockRuntimeClient({ region: "us-east-1" });

// Configuracion de hashing para RUT
const SALT_VERSION = "v1";
const PII_SALT = process.env.PII_SALT ?? "sec-salt-placeholder"; // Inyectar via variable de entorno

/** Genera un hash determinístico de 8 chars para un RUT dado (misma entrada → mismo hash) */
function hashRut(rut) {
  const normalizado = rut.toLowerCase().replace(/[.\s]/g, "");
  return "[RUT_HASH:" +
    crypto.createHmac("sha256", `${SALT_VERSION}:${PII_SALT}`)
      .update(normalizado)
      .digest("hex")
      .slice(0, 8)
      .toUpperCase() +
    "]";
}

/**
 * Limpia PII del texto según la norma interna de anonimización.
 * Categorías cubiertas por regex:
 *   RUT → [RUT_HASH:XXXXXXXX]   (hash determinístico con sal versionada)
 *   Correo electrónico → [EMAIL]
 *   Teléfono → [TELEFONO]
 *   Dirección → [DIRECCION]     (heurística para patrones chilenos)
 *   Número de cuenta → [CUENTA]
 *   Número de medidor → [MEDIDOR]
 *   Coordenadas → [UBICACION]
 *
 *     Nombres propios → requieren NER (ej. Amazon Comprehend DetectPiiEntities).
 *     La regla de privacidad en el prompt instruye al modelo a ignorarlos.
 */
function limpiarPII(texto) {
  let t = texto;

  // 1. RUT  →  [RUT_HASH:XXXXXXXX]
  t = t.replace(/\b\d{1,2}\.?\d{3}\.?\d{3}-[0-9kK]\b/g, (m) => hashRut(m));

  // 2. Correo electrónico  →  [EMAIL]
  t = t.replace(/\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b/g, "[EMAIL]");

  // 3. Teléfono chileno  →  [TELEFONO]
  //    Cubre: +56 9 1234 5678 | +56912345678 | 912345678 | (2)1234-5678 | 22-123-4567
  t = t.replace(
    /(\+?56[\s\-]?)?(\(?\d{1,2}\)?[\s\-]?)?\d{4}[\s.\-]?\d{4}\b/g,
    "[TELEFONO]"
  );

  // 4. Número de cuenta (requiere palabra clave previa)  →  [CUENTA]
  t = t.replace(
    /\b(cuenta|n[°úu]mero\s+de\s+cliente|cliente)\s*[:#°Nº]*\s*\d{4,12}\b/gi,
    "[CUENTA]"
  );

  // 5. Número de medidor (requiere palabra clave previa)  →  [MEDIDOR]
  t = t.replace(
    /\b(medidor|n[°úu]mero\s+de\s+medidor)\s*[:#°Nº]*\s*[A-Z0-9]{4,15}\b/gi,
    "[MEDIDOR]"
  );

  // 6. Coordenadas geográficas (lat, lon)  →  [UBICACION]
  t = t.replace(/-?\d{1,3}\.\d{4,}\s*,\s*-?\d{1,3}\.\d{4,}/g, "[UBICACION]");

  // 7. Direcciones (heurística para Chile)  →  [DIRECCION]
  //    Detecta: "Av. Libertador 1234", "Calle Los Aromos 567 Dpto 2", "Pasaje El Roble 89"
  t = t.replace(
    /\b(av(enida)?\.?|calle|psje\.?|pasaje|villa|población|pobl\.?)\s+[A-ZÁÉÍÓÚÑa-záéíóúñ][A-ZÁÉÍÓÚÑa-záéíóúñ\s]{1,30}\s*#?\s*\d{1,5}(\s*(dpto\.?|depto\.?|departamento|casa|of\.?|oficina)\s*\d{1,4})?\b/gi,
    "[DIRECCION]"
  );

  return t;
}

// Main Code
export const handler = async (event) => {

  if (!event.body) {
    return {
      statusCode: 400,
      body: JSON.stringify({ message: "Se envió un request sin body o body inválido" }),
    };
  }

  const body = JSON.parse(event.body);
  const reclamo = body.reclamo;

// Anonimizacion PII
  const reclamoLimpio = limpiarPII(reclamo);

  // Prompt ajustado según el documento de tipologias
  const prompt = `
  Tengo el siguiente reclamo:
  "${reclamoLimpio}"

  Tarea:
    Eres un analista técnico experto de la SEC (Superintendencia de Electricidad y Combustibles) en Chile. 
    Tu tarea es analizar el reclamo de un usuario y clasificarlo estrictamente en una de las 5 categorías autorizadas.
    Si el relato no se ajusta claramente a ninguna de estas 5 categorías, debes usar la categoría "No Aplica. Enviar a revisión manual con analista SEC".

  REGLA DE DESEMPATE CRÍTICA (Excesivos vs Indebidos):
    Los usuarios suelen quejarse de que la boleta está "muy cara". Para decidir correctamente entre "Cobros Excesivos" y "Cobros Indebidos", debes hacerte esta pregunta: ¿De dónde viene el aumento en el cobro?
    - Si la queja menciona que el consumo de energía (los kWh) subió, que el medidor está marcando de más, o que la casa estuvo vacía y aun así subió el consumo físico -> CLASIFICA COMO "Cobros Excesivos".
    - Si la queja menciona cobros por intereses, multas, arriendo de medidor, reliquidaciones, saldos anteriores de deuda o cobros de servicios que no solicitó (independiente de los kWh) -> CLASIFICA COMO "Cobros Indebidos".

  Categorías autorizadas y criterios técnicos de clasificación:
    1. Cobros Excesivos: El ciudadano señala que su boleta presenta un aumento significativo en el consumo de energía (kWh) respecto de meses anteriores. El foco del reclamo está en que "consumió más de lo habitual" según el registro del medidor (ej. "mi consumo se duplicó", "no he cambiado nada en mi casa y la cuenta subió demasiado"). Excluye reclamos por facturación provisoria, cargos financieros, intereses, arriendos o deuda anterior acumulada.
    
    2. Cobros Indebidos: Cuestionamiento de cargos detallados en la boleta que NO se asocian directamente al consumo medido de energía (kWh). Incluye cobros vinculados al ejercicio financiero (intereses por mora, ajustes, acumulación de saldo anterior, reliquidaciones) o servicios asociados a la distribución (arriendo de medidor, verificación, conexión). Ejemplos: "Me están cobrando intereses que no corresponden", "Me están cobrando arriendo de medidor sin avisar".
    
    3. Consumo No Registrado: Situaciones en las que el consumo eléctrico no es reflejado correctamente en las facturas debido a errores en el registro. Aplica a medidores defectuosos que no registran el consumo real, interrupción en la lectura periódica o consumo estimado incorrecto que no refleja la realidad.
    
    4. Reclamo ERNC: Reclamos relacionados con la implementación o el funcionamiento de sistemas de Energías Renovables No Convencionales (ERNC). Incluye problemas en la conexión de sistemas al suministro eléctrico, falta de beneficios prometidos en proyectos de generación distribuida o demoras en la certificación.
    
    5. Problemas de Lectura: Reclamos relacionados con errores operativos en la lectura de medidores. Incluye lecturas incorrectas que resultan en montos erróneos, problemas en el acceso al medidor por parte del personal de la empresa, o facturación basada en estimaciones cuando era posible realizar lecturas reales.

  Reglas de salida (OBLIGATORIAS):
  - Responder SOLO con un objeto JSON válido.
  - NO usar backticks (\`\`\`) en tu respuesta.
  - NO usar la etiqueta "json".
  - NO agregar texto adicional antes ni después del JSON.
  - El JSON debe ir en una sola línea.
  - Si la categoría es "No Aplica", asigna el valor null a "factor_de_confianza".

  REGLA DE PRIVACIDAD Y ANONIMIZACIÓN:
  - El texto del reclamo puede contener marcadores de anonimización como [EMAIL], [TELEFONO], [RUT_HASH:...], etc. Tratalos como datos normales del contexto.
  - BAJO NINGUNA CIRCUNSTANCIA debes incluir nombres, RUT, direcciones o cualquier dato personal en la "justificacion" del JSON.
  - Si necesitas referirte a la persona en la justificación, utiliza únicamente los términos "el cliente" o "el usuario".

  Formato exacto requerido:
  {"clasificacion": "Nombre de la Categoria", "justificacion": "Texto breve justificando con base en las reglas", "factor_de_confianza": 0.95}
  `;

  try {
    const gemmaResponse = await bedrock.send(new InvokeModelCommand({
      modelId: "google.gemma-3-27b-it",
      contentType: "application/json",
      accept: "application/json",
      body: JSON.stringify({
        messages: [{
          role: "user",
          content: prompt,
        }],
      })
    }));

