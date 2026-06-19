import warnings
import pandas as pd
warnings.filterwarnings("ignore", category=UserWarning)

# ── CONFIGURACIÓN ────────────────────────────────────────────────────────────────
EXCEL_PATH  = "base_recl_LectorBLT_IA.xlsx"
Q_MIN       = 1.0     # Umbral mínimo calidad_score  (pendiente calibración)
MAX_BOLETAS = 24      # Máximo de boletas históricas a considerar

TIPOLOGIA_MAP = {
    "facturación provisoria":       "T1",
    "facturacion provisoria":       "T1",
    "facturación excesiva":         "T2",
    "facturacion excesiva":         "T2",
    "cobros indebidos":             "T3",
    "cobros indebidos — intereses": "T3",
    "cobros indebidos - intereses": "T3",
    "error de lectura":             "T4",
    "problemas de lectura":         "T4",
    "no clasificable":              "T5",
}

TIPOLOGIA_NOMBRE = {
    "T1": "Facturación Provisoria",
    "T2": "Facturación Excesiva",
    "T3": "Cobros Indebidos — Intereses",
    "T4": "Error de Lectura",
    "T5": "No Clasificable",
}

PERFILES_ELEGIBLES = {"C1", "C6", "C8", "C13"}

PERFIL_TIPOLOGIA = {
    "C1": "T1", "C2": "T1", "C3": "T1", "C4": "T1", "C5": "T1",
    "C6": "T2", "C7": "T2", "C8": "T2", "C9": "T2", "C10": "T2",
    "C11": "T3", "C12": "T3",
    "C13": "T4", "C14": "T4",
}

PERFIL_ROL = {
    "C1":  "Primario ELEGIBLE_PARA_DESPACHO — Facturación Provisoria",
    "C2":  "Modificador de C1 — Abono devolución provisoria",
    "C3":  "Alerta / revisión — Señal temprana sin lectura",
    "C4":  "Modificador / revisión — Sin lectura + ciclo irregular",
    "C5":  "Contexto — Cliente especial con estimación",
    "C6":  "Primario ELEGIBLE_PARA_DESPACHO — Facturación Excesiva interanual ≥100%",
    "C7":  "Alerta técnica / revisión — Consumo excede potencia contratada",
    "C8":  "Primario ELEGIBLE_PARA_DESPACHO — Outlier + aumento mensual ≥100%",
    "C9":  "Alerta técnica / revisión — Aumento interanual ≥30% cercano a potencia",
    "C10": "Revisión experta — Cambio de medidor + aumento ≥30%",
    "C11": "Revisión experta — Intereses con tasa alta",
    "C12": "Modificador de C11 — Saldo anterior pendiente",
    "C13": "Primario ELEGIBLE_PARA_DESPACHO — Error de Lectura por ciclo irregular",
    "C14": "Alerta / revisión — Consumo cero o negativo",
}


# ── HELPERS DE LECTURA SEGURA ────────────────────────────────────────────────────

def _get(row: pd.Series, col: str, default=None):
    """Acceso seguro a una celda de pandas Series; retorna default si no existe o es NaN."""
    try:
        val = row[col]
        if pd.isna(val):
            return default
        return val
    except (KeyError, TypeError):
        return default


def _bool(row: pd.Series, col: str) -> bool:
    """Lee una celda booleana de forma segura desde una pandas Series."""
    val = _get(row, col, False)
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return bool(val)
    if isinstance(val, str):
        return val.strip().lower() in ("true", "1", "sí", "si", "yes")
    return False


def _num(row: pd.Series, col: str, default: float = 0.0) -> float:
    """Lee una celda numérica de forma segura desde una pandas Series."""
    val = _get(row, col, default)
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _pct(valor: float) -> str:
    """Formatea un decimal como porcentaje con 1 decimal."""
    return f"{valor * 100:.1f}%"


# ── CARGA Y FILTRADO DEL EXCEL ───────────────────────────────────────────────────

def _cargar_boletas(excel_path: str, ID_REFERENCIA: str) -> tuple:
    """
    Carga el Excel, localiza el cliente por ID de incidente y retorna:
        df_cliente   — hasta MAX_BOLETAS boletas del cliente, ordenadas desc.
        df_reclamada — fila(s) marcadas como 'es boleta reclamada'
    Lanza ValueError si el caso no existe en el Excel.
    """
    df = pd.read_excel(excel_path)

    # Normalizar columna de ID a string
    df["ID de incidente"] = df["ID de incidente"].astype(str).str.strip()
    mask = df["ID de incidente"] == str(ID_REFERENCIA).strip()

    if not mask.any():
        raise ValueError(
            f"Caso {ID_REFERENCIA} no encontrado en '{excel_path}'. "
            f"Verifique que el ID de incidente exista en el Excel."
        )

    cliente_id = df.loc[mask, "cliente"].iloc[0]
    df_cliente = (
        df[df["cliente"] == cliente_id]
        .copy()
        .sort_values("fechaemision", ascending=False)
        .head(MAX_BOLETAS)
    )

    df_reclamada = df_cliente[
        df_cliente["es boleta reclamada"].apply(lambda x: _bool(pd.Series([x]), 0) if not isinstance(x, bool) else x) == True
    ].copy()

    # Fallback más simple y robusto
    try:
        df_reclamada = df_cliente[df_cliente["es boleta reclamada"] == True].copy()
    except Exception:
        df_reclamada = pd.DataFrame()

    return df_cliente, df_reclamada


# ── CONTROLES PREVIOS CP01-CP10 ───────────────────────────────────────────────────

def _controles_previos(df_cliente: pd.DataFrame, df_reclamada: pd.DataFrame) -> dict:
    """
    Evalúa los 10 controles de calidad y exclusión.
    Cada resultado incluye: ok (bool), detalle (descripción técnica),
    valores_observados (dict con los datos reales leídos) y accion (qué ocurre si falla).
    CP10 se inicializa aquí y se actualiza tras evaluar perfiles.
    """
    cp = {}
    n_rec = len(df_reclamada)
    r = df_reclamada.iloc[0] if n_rec > 0 else None

    # ── CP01 — Documento reclamado inequívoco ────────────────────────────────────
    cp["CP01"] = {
        "ok": True,   # Restricción desactivada temporalmente
        "nombre": "Documento reclamado inequívoco",
        "valores_observados": {
            "filas_marcadas_como_reclamada": n_rec,
        },
        "detalle": (
            f"Se encontró exactamente 1 boleta marcada como reclamada."
            if n_rec == 1
            else f"Se encontraron {n_rec} filas marcadas como reclamada "
                 f"(se requiere exactamente 1)."
        ),
        "accion_si_falla": "Sin acción — restricción desactivada temporalmente.",
    }

    # ── CP02 — Tipo documental válido ────────────────────────────────────────────
    if n_rec > 0:
        tipo_doc = str(_get(r, "tipodocumento_texto", "sin dato")).lower()
        ok_cp02  = tipo_doc in ("boleta", "factura")
    else:
        tipo_doc = "sin dato"
        ok_cp02  = False

    cp["CP02"] = {
        "ok": ok_cp02,
        "nombre": "Tipo documental válido",
        "valores_observados": {
            "tipodocumento_texto": tipo_doc,
            "tipos_aceptados": ["boleta", "factura"],
        },
        "detalle": (
            f"Documento reclamado es '{tipo_doc}' — tipo válido para automatización."
            if ok_cp02
            else f"Documento reclamado es '{tipo_doc}' — solo se aceptan boleta o factura. "
                 f"Notas de crédito/débito requieren preprocesamiento."
        ),
        "accion_si_falla": "BLOQUEADO_POR_CONTROL — notas de crédito/débito no son el período principal.",
    }

    # ── CP03 — Deduplicación por doc_key ────────────────────────────────────────
    dup_count = int(df_cliente.duplicated(subset=["doc_key"]).sum())
    docs_dup  = df_cliente[df_cliente.duplicated(subset=["doc_key"], keep=False)]["doc_key"].unique().tolist()

    cp["CP03"] = {
        "ok": True, #dup_count == 0,
        "nombre": "Deduplicación por doc_key",
        "valores_observados": {
            "total_filas_historial": len(df_cliente),
            "filas_duplicadas": dup_count,
            "doc_keys_duplicados": docs_dup[:5],  # máximo 5 para no saturar
        },
        "detalle": (
            f"Sin duplicados en el historial de {len(df_cliente)} boletas."
            if dup_count == 0
            else f"Se detectaron {dup_count} fila(s) duplicada(s) por doc_key "
                 f"en el historial. Esto impide calcular secuencias y promedios correctamente."
        ),
        "accion_si_falla": "BLOQUEADO_POR_CONTROL — duplicados distorsionan indicadores.",
    }

    # ── CP04 — CNR ───────────────────────────────────────────────────────────────
    if n_rec > 0:
        abono_cnr = _bool(r, "abono_cnr")
        monto_cnr = _num(r, "monto_cnr")
        tiene_cnr = abono_cnr or monto_cnr > 0
    else:
        abono_cnr, monto_cnr, tiene_cnr = False, 0.0, False

    cp["CP04"] = {
        "ok": not tiene_cnr,
        "nombre": "Ausencia de CNR",
        "valores_observados": {
            "abono_cnr":  abono_cnr,
            "monto_cnr":  monto_cnr,
        },
        "detalle": (
            "Sin registro CNR en la boleta reclamada — puede continuar flujo estándar."
            if not tiene_cnr
            else f"Registro CNR detectado (abono_cnr={abono_cnr}, monto_cnr=${monto_cnr:,.0f}). "
                 f"Debe procesarse en flujo CNR antes de evaluar Facturación Excesiva."
        ),
        "accion_si_falla": "BLOQUEADO_POR_CONTROL — derivar a flujo CNR (CP04).",
    }

    # ── CP05 — Calidad OCR ───────────────────────────────────────────────────────
    if n_rec > 0:
        score_ocr = _num(r, "calidad_score")
        ok_cp05   = score_ocr >= Q_MIN
    else:
        score_ocr, ok_cp05 = 0.0, False

    cp["CP05"] = {
        "ok": ok_cp05,
        "nombre": "Calidad de reconocimiento óptico (OCR)",
        "valores_observados": {
            "calidad_score":         score_ocr,
            "umbral_minimo_q_min":   Q_MIN,
            "diferencia_vs_umbral":  round(score_ocr - Q_MIN, 4),
        },
        "detalle": (
            f"calidad_score = {score_ocr:.3f} ≥ umbral {Q_MIN} — legibilidad suficiente."
            if ok_cp05
            else f"calidad_score = {score_ocr:.3f} < umbral {Q_MIN} — "
                 f"calidad de extracción insuficiente para automatización."
        ),
        "accion_si_falla": "BLOQUEADO_POR_CONTROL — derivar a experto para revisión manual de boleta.",
    }

    # ── CP06 — Historial suficiente ──────────────────────────────────────────────
    boletas_validas = df_cliente[
        df_cliente["tipodocumento_texto"].str.lower().isin(["boleta", "factura"])
    ]
    n_validas = len(boletas_validas)

    cp["CP06"] = {
        "ok": n_validas >= 3,
        "nombre": "Historial suficiente para comparación",
        "valores_observados": {
            "total_boletas_en_historial":  len(df_cliente),
            "boletas_factura_validas":     n_validas,
            "minimo_requerido":            3,
        },
        "detalle": (
            f"Historial suficiente: {n_validas} boleta(s)/factura(s) válidas disponibles."
            if n_validas >= 3
            else f"Historial insuficiente: solo {n_validas} boleta(s)/factura(s) válidas "
                 f"(mínimo 3). No se pueden calcular comparativos históricos."
        ),
        "accion_si_falla": "DERIVAR — marcar regla como NO_EVALUABLE por falta de historial.",
    }

    # ── CP07 — Variables críticas sin nulos ──────────────────────────────────────
    vars_criticas = [
        "consumo_total", "dias_facturacion", "tipodocumento_texto",
        "fechaemision", "montototal_doc",
    ]
    if n_rec > 0:
        nulos   = [v for v in vars_criticas if _get(r, v) is None]
        valores = {v: _get(r, v) for v in vars_criticas}
    else:
        nulos   = vars_criticas
        valores = {v: None for v in vars_criticas}

    cp["CP07"] = {
        "ok": len(nulos) == 0,
        "nombre": "Variables críticas sin nulos",
        "valores_observados": {
            "variables_revisadas": vars_criticas,
            "variables_nulas":     nulos,
            "valores_leidos":      {k: str(v) for k, v in valores.items()},
        },
        "detalle": (
            "Todas las variables críticas tienen valor — no hay nulos bloqueantes."
            if len(nulos) == 0
            else f"Variables críticas con valor nulo: {nulos}. "
                 f"Sin estos datos no se puede evaluar el caso."
        ),
        "accion_si_falla": "DERIVAR — marcar NO_EVALUABLE por variable(s) crítica(s) ausente(s).",
    }

    # ── CP08 — Medidores ─────────────────────────────────────────────────────────
    if n_rec > 0:
        n_med      = _num(r, "num_medidores_distintos")
        cambio_med = _bool(r, "flag_cambio_medidor")
        nromed     = _get(r, "nromed", "sin dato")
    else:
        n_med, cambio_med, nromed = 0.0, False, "sin dato"

    ok_cp08 = not (n_med > 1 or cambio_med)

    cp["CP08"] = {
        "ok": ok_cp08,
        "nombre": "Unicidad y estabilidad de medidor",
        "valores_observados": {
            "nromed":                  str(nromed),
            "num_medidores_distintos": int(n_med),
            "flag_cambio_medidor":     cambio_med,
        },
        "detalle": (
            f"Medidor estable (nromed={nromed}, sin cambios detectados en el historial)."
            if ok_cp08
            else f"Problema de medidor detectado: {int(n_med)} medidor(es) distinto(s) "
                 f"en el historial, cambio_medidor={cambio_med}. "
                 f"Esto afecta la comparabilidad histórica."
        ),
        "accion_si_falla": "DERIVAR — cambio de medidor invalida comparaciones históricas.",
    }

    # ── CP09 — Ciclo de facturación dentro de rango operacional ─────────────────
    if n_rec > 0:
        dias_fac = _num(r, "dias_facturacion")
        ok_cp09  = 15.0 <= dias_fac <= 95.0
    else:
        dias_fac, ok_cp09 = 0.0, False

    cp["CP09"] = {
        "ok": ok_cp09,
        "nombre": "Ciclo de facturación dentro de rango operacional",
        "valores_observados": {
            "dias_facturacion":    dias_fac,
            "rango_operacional":   "15 a 95 días",
            "fuera_de_rango":      not ok_cp09,
        },
        "detalle": (
            f"Ciclo de {dias_fac:.0f} días dentro del rango operacional (15-95 días)."
            if ok_cp09
            else f"Ciclo de {dias_fac:.0f} días fuera del rango operacional (15-95 días). "
                 f"Puede indicar error en la extracción de datos o período atípico."
        ),
        "accion_si_falla": "DERIVAR — ciclo extremo requiere revisión de extracción.",
    }

    # ── CP10 — Conflicto inter-tipología (se actualiza tras evaluar perfiles) ────
    cp["CP10"] = {
        "ok": True,
        "nombre": "Sin conflicto inter-tipología",
        "valores_observados": {},
        "detalle": "Pendiente — se evalúa tras identificar perfiles activos.",
        "accion_si_falla": "DERIVAR — conflicto entre tipologías requiere resolución experta.",
    }

    return cp


# ── EVALUACIÓN DE PERFILES C1-C14 ────────────────────────────────────────────────

def _evaluar_perfiles(df_cliente: pd.DataFrame, df_reclamada: pd.DataFrame,
                      tipologia: str) -> dict:
    """
    Evalúa los 14 perfiles sobre la boleta reclamada.
    Retorna dict {codigo: {activo, rol, condiciones, detalle}}.
    """
    if len(df_reclamada) == 0:
        return {}

    r = df_reclamada.iloc[0]

    def b(col):  return _bool(r, col)
    def n(col):  return _num(r, col)
    def g(col):  return _get(r, col, "N/D")

    perfiles = {}

    # ── T1 — FACTURACIÓN PROVISORIA ─────────────────────────────────────────────

    # C1: 3+ períodos sin lectura Y el último también es estimado
    c1_tres_sin    = b("flag_tres_o_mas_sin_lectura")
    c1_ultimo_sin  = b("flag_ultimo_sin_lectura")
    c1_activo      = c1_tres_sin and c1_ultimo_sin
    perfiles["C1"] = {
        "activo": c1_activo,
        "rol": PERFIL_ROL["C1"],
        "condiciones": {
            "flag_tres_o_mas_sin_lectura": c1_tres_sin,
            "flag_ultimo_sin_lectura":     c1_ultimo_sin,
            "consecutivo_sin_lectura":     int(n("consecutivo_sin_lectura")),
        },
        "detalle": (
            f"ACTIVO — {int(n('consecutivo_sin_lectura'))} período(s) consecutivos sin lectura real, "
            f"incluido el período reclamado."
            if c1_activo
            else f"Inactivo — tres_o_mas_sin_lectura={c1_tres_sin}, "
                 f"ultimo_sin_lectura={c1_ultimo_sin}."
        ),
    }

    # C2: C1 activo + abono devolución provisoria con monto > 0
    c2_abono  = b("abono_devolucion_provisorio")
    c2_monto  = n("monto_devolucion_provisorio")
    c2_activo = c1_activo and c2_abono and c2_monto > 0
    perfiles["C2"] = {
        "activo": c2_activo,
        "rol": PERFIL_ROL["C2"],
        "condiciones": {
            "C1_activo":                   c1_activo,
            "abono_devolucion_provisorio":  c2_abono,
            "monto_devolucion_provisorio":  c2_monto,
        },
        "detalle": (
            f"ACTIVO — Abono por devolución provisoria de ${c2_monto:,.0f} observado."
            if c2_activo
            else f"Inactivo — C1={c1_activo}, abono={c2_abono}, monto=${c2_monto:,.0f}."
        ),
    }

    # C3: Sin C1; 2 sin lectura + último estimado + desviación >15% vs promedio 3m
    c3_dos_sin   = b("flag_dos_sin_lectura")
    c3_ult_sin   = b("flag_ultimo_sin_lectura")
    c3_desv      = abs(n("variacion_vs_promedio_3m"))
    c3_activo    = not c1_activo and c3_dos_sin and c3_ult_sin and c3_desv > 0.15
    perfiles["C3"] = {
        "activo": c3_activo,
        "rol": PERFIL_ROL["C3"],
        "condiciones": {
            "C1_activo":                  c1_activo,
            "flag_dos_sin_lectura":       c3_dos_sin,
            "flag_ultimo_sin_lectura":    c3_ult_sin,
            "variacion_vs_promedio_3m":   _pct(c3_desv),
            "umbral_desviacion":          "15%",
        },
        "detalle": (
            f"ACTIVO — Señal temprana: 2 períodos sin lectura y desviación del {_pct(c3_desv)} "
            f"respecto al promedio de 3 meses."
            if c3_activo
            else f"Inactivo — C1={c1_activo}, dos_sin_lectura={c3_dos_sin}, "
                 f"desv_3m={_pct(c3_desv)} (umbral >15%)."
        ),
    }

    # C4: Sin lectura + ciclo irregular (modificador, no perfil primario)
    c4_sin_lec  = b("flag_sin_lectura")
    c4_irr      = b("flag_periodo_irregular")
    c4_activo   = c4_sin_lec and c4_irr
    perfiles["C4"] = {
        "activo": c4_activo,
        "rol": PERFIL_ROL["C4"],
        "condiciones": {
            "flag_sin_lectura":      c4_sin_lec,
            "flag_periodo_irregular": c4_irr,
        },
        "detalle": (
            "ACTIVO — Período sin lectura real con ciclo de facturación irregular."
            if c4_activo
            else f"Inactivo — sin_lectura={c4_sin_lec}, periodo_irregular={c4_irr}."
        ),
    }

    # C5: Cliente especial + estimado (contexto, no determinante)
    c5_esp     = b("flag_cliente_especial")
    c5_est     = b("flag_sin_lectura")
    c5_activo  = c5_esp and c5_est
    perfiles["C5"] = {
        "activo": c5_activo,
        "rol": PERFIL_ROL["C5"],
        "condiciones": {
            "flag_cliente_especial": c5_esp,
            "flag_sin_lectura":      c5_est,
        },
        "detalle": (
            "ACTIVO — Cliente marcado como especial con período estimado."
            if c5_activo
            else f"Inactivo — cliente_especial={c5_esp}, estimado={c5_est}."
        ),
    }

    # ── T2 — FACTURACIÓN EXCESIVA ────────────────────────────────────────────────

    lectura_real = not b("flag_sin_lectura")

    # C6: Aumento interanual ≥100% con lectura real
    c6_ia100   = b("flag_aumento_interanual_100")
    c6_var_ia  = n("variacion_interanual")
    c6_activo  = c6_ia100 and lectura_real
    perfiles["C6"] = {
        "activo": c6_activo,
        "rol": PERFIL_ROL["C6"],
        "condiciones": {
            "flag_aumento_interanual_100": c6_ia100,
            "variacion_interanual":        _pct(c6_var_ia),
            "lectura_real":                lectura_real,
        },
        "detalle": (
            f"ACTIVO — Consumo {_pct(c6_var_ia)} superior al año anterior con lectura real."
            if c6_activo
            else f"Inactivo — aumento_interanual_100={c6_ia100} "
                 f"(variación={_pct(c6_var_ia)}), lectura_real={lectura_real}."
        ),
    }

    # C7: Consumo excede potencia contratada (alerta técnica, no elegible)
    c7_exc    = b("flag_consumo_excede_potencia")
    c7_pot    = n("potenciaconectada")
    c7_cons   = n("consumo_total")
    c7_activo = c7_exc and lectura_real
    perfiles["C7"] = {
        "activo": c7_activo,
        "rol": PERFIL_ROL["C7"],
        "condiciones": {
            "flag_consumo_excede_potencia": c7_exc,
            "potenciaconectada_kw":         c7_pot,
            "consumo_total_kwh":            c7_cons,
            "lectura_real":                 lectura_real,
        },
        "detalle": (
            f"ACTIVO — Consumo de {c7_cons:,.0f} kWh supera la referencia de potencia "
            f"contratada ({c7_pot:.1f} kW). Alta prioridad de revisión experta."
            if c7_activo
            else f"Inactivo — consumo_excede_potencia={c7_exc}, lectura_real={lectura_real}."
        ),
    }

    # C8: Outlier estadístico + aumento mensual ≥100% con lectura real
    c8_out    = b("flag_consumo_outlier")
    c8_mens   = b("flag_aumento_consumo_100")
    c8_activo = c8_out and c8_mens and lectura_real
    perfiles["C8"] = {
        "activo": c8_activo,
        "rol": PERFIL_ROL["C8"],
        "condiciones": {
            "flag_consumo_outlier":      c8_out,
            "flag_aumento_consumo_100":  c8_mens,
            "lectura_real":              lectura_real,
            "variacion_mes_anterior":    _pct(n("variacion_mes_anterior")),
        },
        "detalle": (
            f"ACTIVO — Consumo estadísticamente atípico (outlier) con aumento ≥100% "
            f"respecto al mes anterior ({_pct(n('variacion_mes_anterior'))})."
            if c8_activo
            else f"Inactivo — outlier={c8_out}, aumento_mensual_100={c8_mens}, "
                 f"lectura_real={lectura_real}."
        ),
    }

    # C9: Aumento interanual ≥30% + cercano a potencia (excluye C6 y C7)
    c9_ia30   = b("flag_aumento_interanual_30")
    c9_pot    = b("flag_consumo_cercano_potencia")
    c9_activo = c9_ia30 and c9_pot and lectura_real and not c6_activo and not c7_activo
    perfiles["C9"] = {
        "activo": c9_activo,
        "rol": PERFIL_ROL["C9"],
        "condiciones": {
            "flag_aumento_interanual_30":     c9_ia30,
            "flag_consumo_cercano_potencia":  c9_pot,
            "lectura_real":                   lectura_real,
            "C6_activo":                      c6_activo,
            "C7_activo":                      c7_activo,
        },
        "detalle": (
            f"ACTIVO — Aumento interanual ≥30% con consumo cercano a la potencia "
            f"contratada. Requiere verificación de potencia (supuesto S2-S3)."
            if c9_activo
            else f"Inactivo — ia30={c9_ia30}, cercano_potencia={c9_pot}, "
                 f"C6={c6_activo}, C7={c7_activo}."
        ),
    }

    # C10: Cambio de medidor + aumento mensual ≥30% (revisión experta)
    c10_med   = b("flag_cambio_medidor")
    c10_aum   = b("flag_aumento_consumo_30")
    c10_activo = c10_med and c10_aum and lectura_real
    perfiles["C10"] = {
        "activo": c10_activo,
        "rol": PERFIL_ROL["C10"],
        "condiciones": {
            "flag_cambio_medidor":      c10_med,
            "flag_aumento_consumo_30":  c10_aum,
            "lectura_real":             lectura_real,
        },
        "detalle": (
            "ACTIVO — Cambio de medidor coincide con aumento de consumo ≥30%. "
            "El cambio de medidor es hipótesis, no causa acreditada."
            if c10_activo
            else f"Inactivo — cambio_medidor={c10_med}, aumento_30={c10_aum}."
        ),
    }

    # ── T3 — COBROS INDEBIDOS — INTERESES ────────────────────────────────────────

    # C11: Tiene intereses + tasa alta (siempre revisión experta)
    c11_int    = b("tiene_interes")
    c11_tasa_a = b("flag_tasa_interes_alta")
    c11_tasa_v = n("tasa_interes_anual")
    c11_monto  = n("monto_total_intereses")
    c11_activo = c11_int and c11_tasa_a
    perfiles["C11"] = {
        "activo": c11_activo,
        "rol": PERFIL_ROL["C11"],
        "condiciones": {
            "tiene_interes":          c11_int,
            "flag_tasa_interes_alta": c11_tasa_a,
            "tasa_interes_anual":     _pct(c11_tasa_v),
            "monto_total_intereses":  c11_monto,
        },
        "detalle": (
            f"ACTIVO — Intereses cobrados: ${c11_monto:,.0f} a tasa anual {_pct(c11_tasa_v)}. "
            f"Requiere revisión de base y antecedentes de la tasa."
            if c11_activo
            else f"Inactivo — tiene_interes={c11_int}, tasa_alta={c11_tasa_a} "
                 f"(tasa={_pct(c11_tasa_v)})."
        ),
    }

    # C12: C11 activo + saldo anterior pendiente
    c12_saldo  = b("tiene_saldo_anterior")
    c12_monto  = n("monto_saldo_anterior")
    c12_activo = c11_activo and c12_saldo
    perfiles["C12"] = {
        "activo": c12_activo,
        "rol": PERFIL_ROL["C12"],
        "condiciones": {
            "C11_activo":            c11_activo,
            "tiene_saldo_anterior":  c12_saldo,
            "monto_saldo_anterior":  c12_monto,
        },
        "detalle": (
            f"ACTIVO — Saldo anterior de ${c12_monto:,.0f} pendiente. "
            f"Requiere trazabilidad contable del origen del saldo."
            if c12_activo
            else f"Inactivo — C11={c11_activo}, saldo_anterior={c12_saldo} "
                 f"(monto=${c12_monto:,.0f})."
        ),
    }

    # ── T4 — ERROR DE LECTURA ────────────────────────────────────────────────────

    # C13: Período irregular + lectura real + variación ≤30% vs promedio + consumo >0
    c13_irr    = b("flag_periodo_irregular")
    c13_var    = abs(n("variacion_vs_promedio_3m"))
    c13_cons   = n("consumo_total")
    c13_dias   = n("dias_facturacion")
    c13_diario = n("consumo_diario_facturado")
    c13_activo = c13_irr and lectura_real and c13_var <= 0.30 and c13_cons > 0
    perfiles["C13"] = {
        "activo": c13_activo,
        "rol": PERFIL_ROL["C13"],
        "condiciones": {
            "flag_periodo_irregular":     c13_irr,
            "lectura_real":               lectura_real,
            "variacion_vs_promedio_3m":   _pct(c13_var),
            "umbral_variacion":           "≤30%",
            "consumo_total_kwh":          c13_cons,
            "dias_facturacion":           int(c13_dias),
            "consumo_diario_kwh":         round(c13_diario, 3),
        },
        "detalle": (
            f"ACTIVO — Ciclo irregular de {int(c13_dias)} días con lectura real. "
            f"Consumo diario {c13_diario:.2f} kWh/día, variación {_pct(c13_var)} "
            f"respecto al promedio (≤30%: coherente con cambio de duración del ciclo)."
            if c13_activo
            else f"Inactivo — periodo_irregular={c13_irr}, lectura_real={lectura_real}, "
                 f"var_3m={_pct(c13_var)} (umbral ≤30%), consumo={c13_cons:.0f} kWh."
        ),
    }

    # C14: Consumo cero o negativo (siempre derivación)
    c14_cero = b("flag_consumo_cero")
    c14_neg  = b("flag_consumo_negativo")
    c14_activo = c14_cero or c14_neg
    perfiles["C14"] = {
        "activo": c14_activo,
        "rol": PERFIL_ROL["C14"],
        "condiciones": {
            "flag_consumo_cero":     c14_cero,
            "flag_consumo_negativo": c14_neg,
            "consumo_total_kwh":     n("consumo_total"),
        },
        "detalle": (
            f"ACTIVO — Consumo anómalo: cero={c14_cero}, negativo={c14_neg} "
            f"({n('consumo_total'):.3f} kWh). Requiere antecedente externo antes de despacho."
            if c14_activo
            else f"Inactivo — consumo_cero={c14_cero}, consumo_negativo={c14_neg}."
        ),
    }

    return perfiles


# ── PRIORIDAD INTRA-TIPOLOGÍA ────────────────────────────────────────────────────

def _seleccionar_perfil(perfiles: dict, tipologia: str):
    """
    Aplica reglas de prioridad (Sección 9) sobre los perfiles activos
    de la tipología propuesta.
    Retorna (perfil_principal, modificadores, alertas).
    """
    activos = {
        k for k, v in perfiles.items()
        if v["activo"] and PERFIL_TIPOLOGIA.get(k) == tipologia
    }

    perfil_principal = None
    modificadores    = []
    alertas          = []

    if tipologia == "T1":
        if "C1" in activos:
            perfil_principal = "C1"
            for mod in ("C2", "C4", "C5"):
                if mod in activos:
                    modificadores.append(mod)
        elif "C3" in activos:
            alertas.append("C3")

    elif tipologia == "T2":
        # Prioridad: C7 > C6 > C8 > C9 > C10
        for p in ["C7", "C6", "C8", "C9", "C10"]:
            if p in activos:
                if p in ("C7", "C9", "C10"):
                    alertas.append(p)
                else:
                    perfil_principal = p
                    break

    elif tipologia == "T3":
        if "C11" in activos:
            perfil_principal = "C11"
            if "C12" in activos:
                modificadores.append("C12")

    elif tipologia == "T4":
        if "C14" in activos:
            alertas.append("C14")   # C14 prima sobre C13 para derivación
        elif "C13" in activos:
            perfil_principal = "C13"

    return perfil_principal, modificadores, alertas


# ── CONFLICTOS INTER-TIPOLOGÍA (CP10) ────────────────────────────────────────────

def _evaluar_conflictos(perfiles: dict, tipologia: str):
    """Detecta perfiles activos de tipologías distintas a la propuesta."""
    otras = {
        k for k, v in perfiles.items()
        if v["activo"] and PERFIL_TIPOLOGIA.get(k) != tipologia
    }
    if otras:
        detalle_otras = ", ".join(
            f"{k} ({TIPOLOGIA_NOMBRE.get(PERFIL_TIPOLOGIA[k], '?')})"
            for k in sorted(otras)
        )
        return True, (
            f"Señales activas de tipología(s) distintas a {tipologia} "
            f"({TIPOLOGIA_NOMBRE.get(tipologia, '')}): {detalle_otras}."
        )
    return False, "Sin conflictos — todos los perfiles activos corresponden a la tipología propuesta."


# ── ÍNDICE DE SEVERIDAD ──────────────────────────────────────────────────────────

def _calcular_severidad(df_reclamada: pd.DataFrame) -> tuple:
    """
    Calcula el índice de severidad y retorna (score, detalle_familias).
    Regla: dentro de cada familia solo suma la condición de mayor peso.
    """
    if len(df_reclamada) == 0:
        return 0, {}

    r = df_reclamada.iloc[0]

    def b(col): return _bool(r, col)
    def n(col): return _num(r, col)

    score   = 0
    familias = {}

    # Lectura ausente
    consec = n("consecutivo_sin_lectura")
    if consec >= 5:
        pts = 4; desc = f"{int(consec)} períodos consecutivos sin lectura (≥5)"
    elif b("flag_tres_o_mas_sin_lectura"):
        pts = 3; desc = "3 o más períodos sin lectura"
    elif b("flag_dos_sin_lectura"):
        pts = 1; desc = "2 períodos sin lectura"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Lectura ausente"] = {"puntos": pts, "descripcion": desc}

    # Aumento interanual
    var_ia = n("variacion_interanual")
    if var_ia >= 2.00:
        pts = 4; desc = f"variación interanual ≥200% ({_pct(var_ia)})"
    elif b("flag_aumento_interanual_100"):
        pts = 3; desc = f"aumento interanual ≥100% ({_pct(var_ia)})"
    elif b("flag_aumento_interanual_30"):
        pts = 1; desc = f"aumento interanual ≥30% ({_pct(var_ia)})"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Aumento interanual"] = {"puntos": pts, "descripcion": desc}

    # Potencia de referencia
    if b("flag_consumo_excede_potencia"):
        pts = 4; desc = "consumo excede potencia contratada — fuerza revisión experta"
    elif b("flag_consumo_cercano_potencia"):
        pts = 1; desc = "consumo cercano a potencia contratada"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Potencia de referencia"] = {"puntos": pts, "descripcion": desc}

    # Intereses
    if b("flag_tasa_interes_extrema"):
        pts = 4; desc = f"tasa de interés extrema ({_pct(n('tasa_interes_anual'))} anual)"
    elif b("flag_tasa_interes_alta"):
        pts = 2; desc = f"tasa de interés alta ({_pct(n('tasa_interes_anual'))} anual)"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Intereses"] = {"puntos": pts, "descripcion": desc}

    # Outlier estadístico
    if b("flag_consumo_outlier"):
        pts = 2; desc = f"consumo estadísticamente atípico ({n('consumo_total'):,.0f} kWh)"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Outlier estadístico"] = {"puntos": pts, "descripcion": desc}

    # Aumento mensual
    if b("flag_aumento_consumo_100"):
        pts = 2; desc = f"aumento mensual ≥100% ({_pct(n('variacion_mes_anterior'))})"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Aumento mensual"] = {"puntos": pts, "descripcion": desc}

    # Ciclo irregular
    if b("flag_periodo_irregular"):
        pts = 1; desc = f"ciclo irregular ({n('dias_facturacion'):.0f} días)"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Ciclo irregular"] = {"puntos": pts, "descripcion": desc}

    # Cambio de medidor
    if b("flag_cambio_medidor"):
        pts = 1; desc = "cambio de medidor en el historial"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Cambio de medidor"] = {"puntos": pts, "descripcion": desc}

    # Consumo cero o negativo
    if b("flag_consumo_negativo"):
        pts = 2; desc = f"consumo negativo ({n('consumo_total'):.3f} kWh)"
    elif b("flag_consumo_cero"):
        pts = 1; desc = "consumo cero"
    else:
        pts = 0; desc = "sin señal"
    score += pts
    familias["Consumo anómalo"] = {"puntos": pts, "descripcion": desc}

    return score, familias


def _banda_severidad(score: int) -> str:
    if score == 0:   return "Sin señal"
    if score <= 3:   return "Baja"
    if score <= 6:   return "Media"
    if score <= 9:   return "Alta"
    return "Crítica"


# ── TEXTO DE PRERRESOLUCIÓN ───────────────────────────────────────────────────────

def _construir_texto(perfil: str, df_reclamada: pd.DataFrame, modificadores: list) -> str:
    """Plantillas cerradas (Sección 12.1). Solo hechos observables y cálculos reproducibles."""
    if len(df_reclamada) == 0 or perfil is None:
        return ""

    r = df_reclamada.iloc[0]
    def n(col): return _num(r, col)

    periodo  = str(_get(r, "periodo", "sin dato"))
    consumo  = n("consumo_total")
    dias     = int(n("dias_facturacion"))
    var_ia   = n("variacion_interanual")
    var_prom = n("variacion_vs_promedio_3m")
    consec   = int(n("consecutivo_sin_lectura"))
    c_diario = n("consumo_diario_facturado")

    if perfil == "C1":
        texto = (
            f"Del análisis de las boletas asociadas al suministro se observan "
            f"{consec} período(s) consecutivo(s) facturado(s) sin lectura real, "
            f"incluido el período {periodo}. "
            f"La información disponible permite clasificar el reclamo como "
            f"Facturación Provisoria para efectos de su tratamiento."
        )
        if "C2" in modificadores:
            monto_dev = n("monto_devolucion_provisorio")
            if monto_dev > 0:
                texto += (
                    f" Se observa además un abono por devolución provisoria "
                    f"de ${monto_dev:,.0f}."
                )
        texto += (
            " La suficiencia de eventuales ajustes no se presume "
            "a partir de esta clasificación."
        )

    elif perfil in ("C6", "C8"):
        texto = (
            f"La boleta del período {periodo} registra {consumo:,.0f} kWh. "
            f"Este valor presenta una variación de {_pct(var_ia)} respecto del "
            f"período comparable del año anterior y activa criterios internos de "
            f"anomalía cuantitativa. La información disponible permite clasificar "
            f"el reclamo como Facturación Excesiva para efectos de su tratamiento. "
            f"La clasificación identifica una inconsistencia que debe abordarse "
            f"conforme al procedimiento aplicable; no atribuye por sí sola una "
            f"causa técnica específica."
        )

    elif perfil == "C13":
        texto = (
            f"El período reclamado comprende {dias} días, duración que se aparta "
            f"del ciclo habitual. El consumo diario registrado es "
            f"{c_diario:.2f} kWh/día y su variación respecto del promedio reciente "
            f"es {_pct(var_prom)}. La información disponible permite clasificar el "
            f"reclamo como Error de Lectura asociado a la duración del ciclo "
            f"para efectos de su tratamiento."
        )

    else:
        texto = (
            f"El caso del período {periodo} requiere revisión experta. "
            f"No se genera prerresolución automática."
        )

    return texto


# ── RESUMEN EJECUTIVO ─────────────────────────────────────────────────────────────

def _resumen_ejecutivo(
    estado: str, tipologia: str, confianza: float,
    perfil_principal, modificadores: list, alertas: list,
    sev: int, banda_sev: str, familias_sev: dict,
    cp: dict, perfiles: dict,
    causas_derivacion: list, texto_pre: str,
    datos_extractor: dict, df_reclamada: pd.DataFrame,
) -> str:
    """
    Genera un resumen ejecutivo legible en consola con todos los
    controles, perfiles evaluados y decisión final.
    """
    SEP  = "=" * 65
    sep2 = "-" * 65
    lin  = []
    add  = lin.append

    add(SEP)
    add("  RESUMEN EJECUTIVO — LECTOR DE BOLETAS")
    add(SEP)

    # ── Identificación del caso ──────────────────────────────────────────────────
    add("\n[ CASO ]")
    add(f"  ID de incidente : {datos_extractor.get('id', 'N/D')}")
    add(f"  N° referencia   : {datos_extractor.get('n_referencia', 'N/D')}")
    add(f"  Fecha reclamo   : {datos_extractor.get('fecha', 'N/D')}")
    add(f"  Empresa         : {datos_extractor.get('empresa', 'N/D')}")
    add(f"  Categoría SEC   : {datos_extractor.get('cat3', 'N/D')}")

    if len(df_reclamada) > 0:
        r = df_reclamada.iloc[0]
        add(f"  Boleta reclamada: {_get(r, 'doc_key', 'N/D')} "
            f"| Período: {_get(r, 'periodo', 'N/D')} "
            f"| Fecha emisión: {_get(r, 'fechaemision', 'N/D')}")
        add(f"  Consumo total   : {_num(r, 'consumo_total'):,.0f} kWh "
            f"| Días facturación: {int(_num(r, 'dias_facturacion'))}")
        add(f"  Monto total     : ${_num(r, 'montototal_doc'):,.0f}")

    # ── Clasificador textual ─────────────────────────────────────────────────────
    add(f"\n[ CLASIFICADOR TEXTUAL ]")
    add(f"  Tipología propuesta : {tipologia} — {TIPOLOGIA_NOMBRE.get(tipologia, '?')}")
    add(f"  Confianza           : {_pct(confianza)}")

    # ── Controles previos ────────────────────────────────────────────────────────
    add(f"\n[ CONTROLES PREVIOS ]")
    for cod, info in cp.items():
        icono  = "✓" if info["ok"] else "✗"
        estado_cp = "OK" if info["ok"] else "FALLA"
        add(f"  {icono} {cod} — {info.get('nombre', '')} [{estado_cp}]")
        add(f"       {info['detalle']}")
        if not info["ok"]:
            add(f"       ⚠ Acción: {info.get('accion_si_falla', 'Ver especificación.')}")

    # ── Perfiles evaluados ────────────────────────────────────────────────────────
    add(f"\n[ PERFILES C1-C14 EVALUADOS ]")
    perfiles_activos = [k for k, v in perfiles.items() if v["activo"]]
    for cod in ["C1","C2","C3","C4","C5","C6","C7","C8","C9","C10","C11","C12","C13","C14"]:
        if cod not in perfiles:
            continue
        info  = perfiles[cod]
        icono = "●" if info["activo"] else "○"
        add(f"  {icono} {cod} {'[ACTIVO]' if info['activo'] else '[inactivo]'} — {info['rol']}")
        add(f"       {info['detalle']}")

    # ── Conflictos inter-tipología ────────────────────────────────────────────────
    add(f"\n[ PERFILES ACTIVOS TOTALES ]")
    if perfiles_activos:
        for p in perfiles_activos:
            rol_tp = TIPOLOGIA_NOMBRE.get(PERFIL_TIPOLOGIA.get(p, ""), "?")
            add(f"  • {p} ({rol_tp})")
    else:
        add("  (ninguno)")

    # ── Severidad ────────────────────────────────────────────────────────────────
    add(f"\n[ ÍNDICE DE SEVERIDAD ]")
    add(f"  Score total : {sev} puntos → Banda: {banda_sev}")
    for familia, detalle in familias_sev.items():
        if detalle["puntos"] > 0:
            add(f"  + {detalle['puntos']:2d} pt — {familia}: {detalle['descripcion']}")

    # ── Selección de perfil ───────────────────────────────────────────────────────
    add(f"\n[ SELECCIÓN DE PERFIL ]")
    add(f"  Perfil principal : {perfil_principal or '(ninguno)'}")
    add(f"  Modificadores    : {modificadores or '(ninguno)'}")
    add(f"  Alertas activas  : {alertas or '(ninguno)'}")

    # ── Decisión final ────────────────────────────────────────────────────────────
    add(f"\n[ DECISIÓN FINAL ]")
    add(f"  Estado : {estado}")
    if causas_derivacion:
        add(f"  Causas de derivación / bloqueo:")
        for causa in causas_derivacion:
            add(f"    • {causa}")

    if texto_pre:
        add(f"\n[ PRERRESOLUCIÓN AUTOMÁTICA ]")
        for linea in texto_pre.split(". "):
            linea = linea.strip()
            if linea:
                add(f"  {linea}.")

    add(f"\n{SEP}\n")
    return "\n".join(lin)


# ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────────────────────────

def analizar(
    datos_extractor: dict,
    clasificacion_textual: dict,
    excel_path: str = EXCEL_PATH,
) -> dict:
    """
    Ejecuta el lector de boletas completo.

    Parameters
    ----------
    datos_extractor       : dict  — resultado de Extractor.extraer()
    clasificacion_textual : dict  — resultado de Clasificador.clasificar()
    excel_path            : str   — ruta al Excel de boletas

    Returns
    -------
    dict con: estado, tipologia, perfil_principal, modificadores, alertas,
              severidad, banda_severidad, familias_severidad, controles_previos,
              perfiles, texto_prerresolucion, reporte_operacional,
              resumen_ejecutivo, causal_derivacion
    """
    ID_REFERENCIA = str(datos_extractor.get("id", "")).strip()

    # ── 1. Cargar boletas ────────────────────────────────────────────────────────
    try:
        df_cliente, df_reclamada = _cargar_boletas(excel_path, ID_REFERENCIA)
    except ValueError as e:
        msg = str(e)
        print(f"[LectorBoletas] NO_EVALUABLE — {msg}")
        return {
            "estado": "NO_EVALUABLE", "tipologia": None,
            "perfil_principal": None, "modificadores": [], "alertas": [],
            "severidad": 0, "banda_severidad": "Sin señal", "familias_severidad": {},
            "controles_previos": {}, "perfiles": {},
            "texto_prerresolucion": "", "reporte_operacional": {},
            "resumen_ejecutivo": f"NO_EVALUABLE — {msg}",
            "causal_derivacion": msg,
        }

    # ── 2. Normalizar tipología ──────────────────────────────────────────────────
    tipo_raw  = str(clasificacion_textual.get("clasificacion", "")).strip().lower()
    confianza = _num(pd.Series({"v": clasificacion_textual.get("factor_de_confianza", 0)}), "v")
    tipologia = TIPOLOGIA_MAP.get(tipo_raw, "T5")

    if tipologia == "T5":
        sev, familias = _calcular_severidad(df_reclamada)
        causal = (
            f"Clasificador textual emitió tipología no reconocida o T5 "
            f"(entrada: '{tipo_raw}') — derivación automática sin evaluación de perfiles."
        )
        resumen = _resumen_ejecutivo(
            "CLASIFICADO_PARA_REVISION", "T5", confianza,
            None, [], [], sev, _banda_severidad(sev), familias,
            {}, {}, [causal], "", datos_extractor, df_reclamada,
        )
        print(resumen)
        return {
            "estado": "CLASIFICADO_PARA_REVISION", "tipologia": "T5",
            "perfil_principal": None, "modificadores": [], "alertas": [],
            "severidad": sev, "banda_severidad": _banda_severidad(sev),
            "familias_severidad": familias,
            "controles_previos": {}, "perfiles": {},
            "texto_prerresolucion": "", "reporte_operacional": {},
            "resumen_ejecutivo": resumen, "causal_derivacion": causal,
        }

    # ── 3. Controles previos CP01-CP09 ───────────────────────────────────────────
    cp = _controles_previos(df_cliente, df_reclamada)
    causas_derivacion = []

    # Bloqueantes: detienen el flujo inmediatamente
    for cod in ("CP02", "CP03", "CP04", "CP05"):  # CP01 desactivado
        if not cp[cod]["ok"]:
            sev, familias = _calcular_severidad(df_reclamada)
            causal = f"{cod} ({cp[cod]['nombre']}): {cp[cod]['detalle']}"
            resumen = _resumen_ejecutivo(
                "BLOQUEADO_POR_CONTROL", tipologia, confianza,
                None, [], [], sev, _banda_severidad(sev), familias,
                cp, {}, [causal], "", datos_extractor, df_reclamada,
            )
            print(resumen)
            return {
                "estado": "BLOQUEADO_POR_CONTROL", "tipologia": tipologia,
                "perfil_principal": None, "modificadores": [], "alertas": [],
                "severidad": sev, "banda_severidad": _banda_severidad(sev),
                "familias_severidad": familias,
                "controles_previos": cp, "perfiles": {},
                "texto_prerresolucion": "", "reporte_operacional": {},
                "resumen_ejecutivo": resumen, "causal_derivacion": causal,
            }

    # Derivantes: acumulan causas pero no detienen la evaluación de perfiles
    for cod in ("CP06", "CP07", "CP08", "CP09"):
        if not cp[cod]["ok"]:
            causas_derivacion.append(
                f"{cod} ({cp[cod]['nombre']}): {cp[cod]['detalle']}"
            )

    # ── 4. Evaluación de perfiles ────────────────────────────────────────────────
    perfiles = _evaluar_perfiles(df_cliente, df_reclamada, tipologia)

    # ── 5. Conflictos inter-tipología → CP10 ────────────────────────────────────
    hay_conflicto, detalle_conflicto = _evaluar_conflictos(perfiles, tipologia)
    perfiles_activos_otros = {
        k for k, v in perfiles.items()
        if v["activo"] and PERFIL_TIPOLOGIA.get(k) != tipologia
    }
    cp["CP10"] = {
        "ok":      not hay_conflicto,
        "nombre":  "Sin conflicto inter-tipología",
        "valores_observados": {
            "tipologia_propuesta":          tipologia,
            "perfiles_activos_otras_tipol": sorted(perfiles_activos_otros),
        },
        "detalle":             detalle_conflicto,
        "accion_si_falla":     "DERIVAR — conflicto entre tipologías requiere resolución experta.",
    }
    if hay_conflicto:
        causas_derivacion.append(f"CP10: {detalle_conflicto}")

    # ── 6. Selección de perfil ───────────────────────────────────────────────────
    perfil_principal, modificadores, alertas = _seleccionar_perfil(perfiles, tipologia)

    # ── 7. Verificar señal compatible ────────────────────────────────────────────
    activos_tipologia = {
        k for k, v in perfiles.items()
        if v["activo"] and PERFIL_TIPOLOGIA.get(k) == tipologia
    }
    if not activos_tipologia:
        sev, familias = _calcular_severidad(df_reclamada)
        causal = (
            f"SIN_SENAL_COMPATIBLE — ningún perfil de {tipologia} "
            f"({TIPOLOGIA_NOMBRE.get(tipologia, '')}) se activó con los datos disponibles. "
            f"Las boletas no muestran señales consistentes con la tipología propuesta."
        )
        resumen = _resumen_ejecutivo(
            "SIN_SENAL_COMPATIBLE", tipologia, confianza,
            None, [], [], sev, _banda_severidad(sev), familias,
            cp, perfiles, [causal], "", datos_extractor, df_reclamada,
        )
        print(resumen)
        return {
            "estado": "SIN_SENAL_COMPATIBLE", "tipologia": tipologia,
            "perfil_principal": None, "modificadores": [], "alertas": [],
            "severidad": sev, "banda_severidad": _banda_severidad(sev),
            "familias_severidad": familias,
            "controles_previos": cp, "perfiles": perfiles,
            "texto_prerresolucion": "", "reporte_operacional": {},
            "resumen_ejecutivo": resumen, "causal_derivacion": causal,
        }

    # ── 8. Severidad ─────────────────────────────────────────────────────────────
    sev, familias = _calcular_severidad(df_reclamada)
    banda_sev = _banda_severidad(sev)

    # ── 9. Estado final ──────────────────────────────────────────────────────────
    if causas_derivacion:
        estado = "CLASIFICADO_PARA_REVISION"
    elif perfil_principal is None:
        estado = "CLASIFICADO_PARA_REVISION"
        causas_derivacion.append(
            "No se identificó perfil primario elegible activo para la tipología propuesta."
        )
    elif perfil_principal not in PERFILES_ELEGIBLES:
        estado = "CLASIFICADO_PARA_REVISION"
        causas_derivacion.append(
            f"Perfil {perfil_principal} ({PERFIL_ROL.get(perfil_principal, '')}) "
            f"no está habilitado para despacho automático en la versión actual."
        )
    elif alertas:
        estado = "CLASIFICADO_PARA_REVISION"
        alertas_desc = ", ".join(
            f"{a} ({PERFIL_ROL.get(a, '')})" for a in alertas
        )
        causas_derivacion.append(
            f"Alertas activas impiden despacho automático: {alertas_desc}."
        )
    else:
        estado = "VALIDADO_PARA_DESPACHO"

    # ── 10. Texto de prerresolución ───────────────────────────────────────────────
    texto_pre = ""
    if estado == "VALIDADO_PARA_DESPACHO":
        texto_pre = _construir_texto(perfil_principal, df_reclamada, modificadores)

    # ── 11. Reporte operacional (para auditoría) ──────────────────────────────────
    reporte_op: dict = {
        "ID_REFERENCIA":              ID_REFERENCIA,
        "empresa":              datos_extractor.get("empresa"),
        "fecha_reclamo":        datos_extractor.get("fecha"),
        "tipologia_propuesta":  tipologia,
        "nombre_tipologia":     TIPOLOGIA_NOMBRE.get(tipologia),
        "confianza_textual":    confianza,
        "perfil_principal":     perfil_principal,
        "modificadores":        modificadores,
        "alertas":              alertas,
        "severidad":            sev,
        "banda_severidad":      banda_sev,
        "controles_previos":    {k: {"ok": v["ok"], "detalle": v["detalle"]} for k, v in cp.items()},
        "perfiles_activos":     [k for k, v in perfiles.items() if v["activo"]],
        "causal_derivacion":    "; ".join(causas_derivacion) if causas_derivacion else "",
    }
    if len(df_reclamada) > 0:
        r = df_reclamada.iloc[0]
        reporte_op.update({
            "doc_key":                  str(_get(r, "doc_key", "")),
            "periodo":                  str(_get(r, "periodo", "")),
            "consumo_total_kwh":        _num(r, "consumo_total"),
            "dias_facturacion":         int(_num(r, "dias_facturacion")),
            "montototal_doc":           _num(r, "montototal_doc"),
            "variacion_interanual":     _num(r, "variacion_interanual"),
            "variacion_vs_promedio_3m": _num(r, "variacion_vs_promedio_3m"),
            "consecutivo_sin_lectura":  int(_num(r, "consecutivo_sin_lectura")),
            "calidad_score":            _num(r, "calidad_score"),
        })

    # ── 12. Resumen ejecutivo ─────────────────────────────────────────────────────
    resumen = _resumen_ejecutivo(
        estado, tipologia, confianza,
        perfil_principal, modificadores, alertas,
        sev, banda_sev, familias,
        cp, perfiles, causas_derivacion, texto_pre,
        datos_extractor, df_reclamada,
    )
    print(resumen)

    return {
        "estado":               estado,
        "tipologia":            tipologia,
        "perfil_principal":     perfil_principal,
        "modificadores":        modificadores,
        "alertas":              alertas,
        "severidad":            sev,
        "banda_severidad":      banda_sev,
        "familias_severidad":   familias,
        "controles_previos":    cp,
        "perfiles":             perfiles,
        "texto_prerresolucion": texto_pre,
        "reporte_operacional":  reporte_op,
        "resumen_ejecutivo":    resumen,
        "causal_derivacion":    "; ".join(causas_derivacion) if causas_derivacion else "",
    }


# ── EJECUCIÓN DIRECTA ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import json
    from Extractor import extraer, ID_REFERENCIA, ID_PENDIENTES
    from Clasificador import clasificar

    datos_caso    = extraer(ID_REFERENCIA=ID_REFERENCIA, id_pendientes=ID_PENDIENTES)
    clasificacion = clasificar(datos_caso)
    reporte       = analizar(datos_caso, clasificacion)