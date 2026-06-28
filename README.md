# Sistema de Clasificación y Prerresolución de Reclamos Eléctricos — SEC

Pipeline automatizado que extrae un reclamo desde la base de datos de la SEC, lo clasifica mediante un modelo de IA y evalúa si puede resolverse automáticamente contrastando boletas históricas, según la especificación funcional v2.2 (Equipo IEE3112, PUC).

```
Extractor.py  →  Clasificador.py  →  LectorBoletas.py
   (SEC API)      (IA / AWS Lambda)     (Excel boletas)
```

Incluye una interfaz gráfica de escritorio (`main.py`, PyQt5) para ejecutar el flujo completo sin usar la consola.

---

## Tabla de contenidos

- [Sistema de Clasificación y Prerresolución de Reclamos Eléctricos — SEC](#sistema-de-clasificación-y-prerresolución-de-reclamos-eléctricos--sec)
  - [Tabla de contenidos](#tabla-de-contenidos)
  - [Requisitos](#requisitos)
    - [Dependencias](#dependencias)
  - [Instalación](#instalación)
  - [Estructura del proyecto](#estructura-del-proyecto)
  - [Configuración de credenciales](#configuración-de-credenciales)
  - [Uso con interfaz gráfica](#uso-con-interfaz-gráfica)
  - [Módulo `Extractor.py`](#módulo-extractorpy)
  - [Módulo `Clasificador.py`](#módulo-clasificadorpy)
  - [Módulo `LectorBoletas.py`](#módulo-lectorboletaspy)
    - [Configuración interna](#configuración-interna)
  - [Estados de salida del lector](#estados-de-salida-del-lector)
  - [Perfiles C1–C14](#perfiles-c1c14)
  - [Controles previos CP01–CP10](#controles-previos-cp01cp10)
  - [Formato del Excel de boletas](#formato-del-excel-de-boletas)
  - [Solución de problemas](#solución-de-problemas)

---

## Requisitos

- Python 3.10 o superior
- Conexión a internet (API de SEC y API de clasificación)
- Excel de boletas históricas (`base_recl_LectorBLT_IA.xlsx` o equivalente)

### Dependencias

```bash
pip install requests pandas openpyxl PyQt5
```

| Paquete | Uso |
|---|---|
| `requests` | Llamadas HTTP a la API de SEC y a la API de clasificación |
| `pandas` | Lectura y procesamiento del Excel de boletas |
| `openpyxl` | Motor de lectura de archivos `.xlsx` usado por pandas |
| `PyQt5` | Interfaz gráfica de escritorio (`gui_app.py`) |

---

## Instalación

```bash
git clone <repositorio>
cd <carpeta-del-proyecto>
pip install requests pandas openpyxl PyQt5
```

Coloca el Excel de boletas en la misma carpeta que los módulos, o ajusta la ruta en `LectorBoletas.py` (constante `EXCEL_PATH`).

---

## Estructura del proyecto

```
proyecto/
├── Extractor.py        # Descarga el caso desde la API de SEC
├── Clasificador.py     # Clasifica el texto del reclamo vía API de IA
├── LectorBoletas.py    # Motor determinístico de perfiles C1-C14
├── main.py              # Flujo principal y GUI
└── base_recl_LectorBLT_IA.xlsx   # Excel de boletas históricas (no incluido en el repo)
```

---

## Configuración de credenciales

Antes de ejecutar, completa las credenciales en los siguientes archivos:

**`Extractor.py`**
```python
SEC_USERNAME = "ProyectoUC"
SEC_PASSWORD = "TU_CONTRASEÑA"     # Reemplazar
```

**`Clasificador.py`**
```python
CLASIFICADOR_API_KEY = "TU_API_KEY"     # Reemplazar
```
---
## Uso con interfaz gráfica

```bash
python main.py
```

La ventana permite:

- Ingresar el **N° de referencia SEC** del caso.
- Ingresar el **ID de incidente** correspondiente en el Excel de boletas (con un valor por defecto precargado).
- Ejecutar el pipeline completo con un botón, viendo el avance por etapas (Extractor → Clasificador → Lector de Boletas) en tarjetas de estado y una barra de progreso.
- Revisar el log de ejecución en una consola integrada, con jerarquía visual (encabezados, controles aprobados/fallidos, perfiles activos/inactivos).
- Ver el resultado final resumido: estado de despacho, tipología, perfil, severidad y prerresolución generada (si corresponde).

El pipeline corre en un hilo separado (`QThread`) para no bloquear la interfaz mientras se realizan las consultas HTTP y la lectura del Excel.

---

## Módulo `Extractor.py`

Descarga el reporte de casos pendientes desde la API de SEC (`analyticsReportResults`) y localiza la fila correspondiente al N° de referencia indicado.

```python
def extraer(ID_REFERENCIA: str = ID_REFERENCIA, id_pendientes: int = ID_PENDIENTES) -> dict
```

**Retorna** un diccionario con:

| Campo | Descripción |
|---|---|
| `id` | ID de incidente SEC (usado por `LectorBoletas` para ubicar al cliente en el Excel) |
| `n_referencia` | N° de referencia del caso |
| `fecha` | Fecha de creación |
| `estado` | Estado actual del caso en SEC |
| `cat3` / `cat4` | Categorías de clasificación SEC |
| `empresa` | Empresa eléctrica reclamada |
| `peticion` | Petición concreta del reclamante |
| `texto` | Texto completo del reclamo |
| `respuesta_empresa` | Primera respuesta de la empresa (si existe) |

**Autenticación:** HTTP Basic Auth (`SEC_USERNAME` / `SEC_PASSWORD`).

---

## Módulo `Clasificador.py`

Envía el texto del reclamo a una API de clasificación basada en IA (AWS Lambda) y obtiene una tipología propuesta con nivel de confianza.

```python
def clasificar(datos_caso: dict) -> dict
```

**Entrada:** el dict retornado por `Extractor.extraer()` (usa la clave `"texto"`).

**Retorna:**

| Campo | Descripción |
|---|---|
| `clasificacion` | Tipología textual propuesta (ej. `"Facturación Provisoria"`) |
| `factor_de_confianza` | Confianza del modelo, entre 0 y 1 |
| `justificacion` | Explicación textual de la clasificación |

**Autenticación:** header `x-api-key` (`CLASIFICADOR_API_KEY`).

---

## Módulo `LectorBoletas.py`

Motor determinístico que contrasta la tipología propuesta por el clasificador con las boletas históricas del cliente, según la especificación funcional v2.2 (Secciones 2, 6, 7, 8, 9, 10 y 11 del documento de referencia).

```python
def analizar(datos_extractor: dict, clasificacion_textual: dict, excel_path: str = EXCEL_PATH) -> dict
```

**Flujo interno:**

1. Carga el Excel y localiza hasta `MAX_BOLETAS` (24) boletas del cliente asociado al `id` del caso.
2. Normaliza la tipología textual del clasificador a un código interno T1–T5 (`TIPOLOGIA_MAP`).
3. Evalúa los controles previos CP01–CP10.
4. Evalúa los 14 perfiles C1–C14 aplicables a la tipología propuesta.
5. Aplica las reglas de prioridad intra e inter-tipología para seleccionar el perfil principal.
6. Calcula el índice de severidad.
7. Determina el estado final de despacho.
8. Si corresponde, construye el texto de prerresolución con plantilla cerrada.
9. Genera un resumen ejecutivo legible con el detalle completo de la decisión.

**Retorna** un diccionario con:

| Campo | Descripción |
|---|---|
| `estado` | Uno de los 5 [estados de salida](#estados-de-salida-del-lector) |
| `tipologia` | Código T1–T5 |
| `perfil_principal` | Perfil C1–C14 seleccionado, o `None` |
| `modificadores` | Lista de perfiles modificadores activos |
| `alertas` | Lista de perfiles de alerta activos |
| `severidad` | Puntaje numérico de severidad |
| `banda_severidad` | `Sin señal` / `Baja` / `Media` / `Alta` / `Crítica` |
| `familias_severidad` | Detalle del puntaje por familia de riesgo |
| `controles_previos` | Resultado detallado de CP01–CP10 |
| `perfiles` | Resultado detallado de C1–C14 |
| `texto_prerresolucion` | Texto generado (vacío si no hay despacho automático) |
| `reporte_operacional` | Variables clave para auditoría interna |
| `resumen_ejecutivo` | Texto formateado para consola/log |
| `causal_derivacion` | Motivo de derivación, si corresponde |

### Configuración interna

```python
EXCEL_PATH  = "base_recl_LectorBLT_IA.xlsx"   # Ruta al Excel de boletas
Q_MIN       = 1.0                              # Umbral mínimo de calidad_score
MAX_BOLETAS = 24                               # Máximo de boletas históricas a considerar
```

---

## Estados de salida del lector

Definidos en el Anexo B de la especificación funcional:

| Estado | Significado |
|---|---|
| `VALIDADO_PARA_DESPACHO` | El lector confirma que la tipología inicial puede resolverse mediante un perfil elegible; el expediente puede pasar al módulo de autorización |
| `CLASIFICADO_PARA_REVISION` | Existe un perfil aplicable, pero el expediente requiere revisión experta |
| `NO_EVALUABLE` | Faltan datos o historial para ejecutar una regla |
| `BLOQUEADO_POR_CONTROL` | Existe una excepción de calidad o exclusión (CP01–CP05) |
| `SIN_SENAL_COMPATIBLE` | Ningún perfil de la tipología propuesta se activó con los datos disponibles — salvaguarda contra errores de clasificación textual |

---

## Perfiles C1–C14

| Perfil | Tipología | Rol | Elegible para despacho |
|---|---|---|---|
| C1 | T1 — Facturación Provisoria | Primario | ✅ |
| C2 | T1 | Modificador de C1 | — |
| C3 | T1 | Alerta / revisión | ❌ |
| C4 | T1 | Modificador / revisión | — |
| C5 | T1 | Contexto | — |
| C6 | T2 — Facturación Excesiva | Primario | ✅ |
| C7 | T2 | Alerta técnica / revisión | ❌ |
| C8 | T2 | Primario | ✅ |
| C9 | T2 | Alerta técnica / revisión | ❌ |
| C10 | T2 | Revisión experta | ❌ |
| C11 | T3 — Cobros Indebidos / Intereses | Revisión experta | ❌ |
| C12 | T3 | Modificador de C11 | — |
| C13 | T4 — Error de Lectura | Primario | ✅ |
| C14 | T4 | Alerta / revisión | ❌ |

Solo **C1, C6, C8 y C13** pueden derivar en despacho automático (`PERFILES_ELEGIBLES`); el resto siempre deriva a revisión experta cuando se activa.

---

## Controles previos CP01–CP10

| Código | Control | Bloqueante |
|---|---|---|
| CP01 | Documento reclamado inequívoco | ✅ *(desactivable, ver nota)* |
| CP02 | Tipo documental válido (boleta/factura) | ✅ |
| CP03 | Deduplicación por `doc_key` | ✅ |
| CP04 | Ausencia de CNR | ✅ |
| CP05 | Calidad de OCR (`calidad_score >= Q_MIN`) | ✅ |
| CP06 | Historial suficiente (≥3 boletas) | Deriva, no bloquea |
| CP07 | Variables críticas sin nulos | Deriva, no bloquea |
| CP08 | Unicidad y estabilidad de medidor | Deriva, no bloquea |
| CP09 | Ciclo de facturación en rango (15–95 días) | Deriva, no bloquea |
| CP10 | Sin conflicto inter-tipología | Deriva, no bloquea |

> **Nota:** CP01 puede estar temporalmente desactivado en el código (`"ok": True` fijo) para pruebas. Para reactivarlo, restaurar la condición `n_rec == 1` y volver a incluir `"CP01"` en la tupla de controles bloqueantes dentro de `analizar()`.

---

## Formato del Excel de boletas

`LectorBoletas.py` espera un archivo `.xlsx` con (al menos) las siguientes columnas:

**Identificación**
`ID de incidente`, `cliente`, `doc_key`, `periodo`, `fechaemision`, `tipodocumento_texto`, `es boleta reclamada`

**Consumo**
`consumo_total`, `consumo_diario_facturado`, `consumo_mes_anterior`, `consumo_mes_anio_anterior`

**Variaciones**
`variacion_mes_anterior`, `variacion_interanual`, `variacion_vs_promedio_3m`

**Lectura**
`flag_sin_lectura`, `flag_dos_sin_lectura`, `flag_tres_o_mas_sin_lectura`, `flag_ultimo_sin_lectura`, `consecutivo_sin_lectura`

**Ciclo**
`dias_facturacion`, `flag_periodo_irregular`

**Potencia**
`potenciaconectada`, `flag_consumo_cercano_potencia`, `flag_consumo_excede_potencia`

**Medidor**
`nromed`, `num_medidores_distintos`, `flag_cambio_medidor`

**Cargos**
`tiene_interes`, `flag_tasa_interes_alta`, `flag_tasa_interes_extrema`, `tasa_interes_anual`, `monto_total_intereses`, `tiene_saldo_anterior`, `monto_saldo_anterior`

**Ajustes**
`abono_devolucion_provisorio`, `monto_devolucion_provisorio`, `abono_cnr`, `monto_cnr`

**Calidad**
`calidad_score`

**Montos**
`montototal_doc`

**Flags de aumento**
`flag_aumento_interanual_30`, `flag_aumento_interanual_100`, `flag_aumento_consumo_30`, `flag_aumento_consumo_100`, `flag_consumo_outlier`

**Consumo anómalo**
`flag_consumo_cero`, `flag_consumo_negativo`

**Contexto**
`flag_cliente_especial`

---

## Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| `403 Client Error: Forbidden` en `Extractor.py` | Credenciales SEC inválidas o expiradas | Verificar `SEC_USERNAME` / `SEC_PASSWORD` |
| `400 Client Error: Bad Request` en `Extractor.py` | Body del request mal formado | Confirmar que solo se envía `{"id": id_pendientes}` |
| `502 Bad Gateway` en `Clasificador.py` | Payload no coincide con lo que espera el Lambda | Confirmar que se envía `{"reclamo": texto}` como string plano |
| `ValueError: Caso ... no encontrado` en `LectorBoletas.py` | El `ID de incidente` no existe en el Excel | Verificar que el ID coincide exactamente con una fila del Excel |
| Carga del Excel muy lenta (~30–40 s) | Archivo grande (decenas de miles de filas) con motor `openpyxl` | Comportamiento esperado; considerar cachear la lectura si se repite muy seguido |
| `ImportError: cannot import name 'extraer' from partially initialized module` | Import circular entre módulos | Asegurar que `Clasificador.py` y `LectorBoletas.py` no importen `Extractor` directamente; solo `main.py` / `gui_app.py` deben orquestar los tres |

---