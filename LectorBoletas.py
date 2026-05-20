import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Optional


class LectorBoletas:
    """
    Lector y analizador del historial de boletas de un cliente eléctrico (CGE/SEC).

    Extrae registros por cliente y calcula métricas orientadas a la resolución
    de reclamos: discrepancias de medición, anomalías de consumo, irregularidades
    en períodos de facturación y resumen del historial de cobros.

    Flujo típico
    ────────────
    lector = LectorBoletas("base_recl_LectorBLT_IA.xlsx")
    lector.cargar_cliente(1130340)

    resumen   = lector.resumen_cliente()
    consumo   = lector.metricas_consumo()
    discrepan = lector.discrepancias_medicion()
    anomalias = lector.detectar_anomalias()
    contexto  = lector.generar_contexto_reclamo()
    """

    # ── Mapeo de columnas del Excel ──────────────────────────────────────────
    COL_CLIENTE        = "cliente"
    COL_EMPRESA        = "empresa"
    COL_NUMDOC         = "numdocumento"
    COL_FECHA_EMISION  = "fechaemision"
    COL_TIPO_DOC       = "tipodocumento_texto"
    COL_CONSUMO        = "consumo_total"
    COL_DIAS_FACT      = "dias_facturacion"
    COL_DIAS_LECTURA   = "dias_periodo_lectura"
    COL_MONTO_ELEC     = "monto_electrico"
    COL_MONTO_TOTAL    = "montototal_doc"
    COL_PRECIO_KWH     = "precio_kwh"
    COL_CONSUMO_DIARIO = "consumo_diario_facturado"
    COL_CONSUMO_TEOR   = "consumo_teorico_maximo"
    COL_CONSUMO_PROY   = "consumo_diario_proyectado_mes"
    COL_POTENCIA       = "potenciaconectada"
    COL_TARIFA         = "tipo_tarifa"
    COL_SECTOR         = "sector_tarifario"
    COL_FECHA_LEC_ACT  = "fechalecactual"
    COL_FECHA_LEC_ANT  = "fechalecant"
    COL_MONTO_ARRIENDO = "monto_arriendo_medidor"
    COL_MONTO_DEVOL    = "monto_devolucion_provisorio"
    COL_MONTO_COMP     = "monto_compensacion"
    COL_TIPO_COMP      = "tipos_compensacion"
    COL_MONTO_SUBS     = "monto_subsidio"
    COL_MONTO_PAND     = "monto_subsidio_pandemia"
    COL_MONTO_CUOTA    = "monto_cuota_pandemia"

        # ── Flags booleanos ──────────────────────────────────────────────────────────
    FLAG_SIN_LECTURA     = "flag_sin_lectura"
    FLAG_CONS_CERO       = "flag_consumo_cero"
    FLAG_CONS_NEG        = "flag_consumo_negativo"
    FLAG_OUTLIER         = "flag_consumo_outlier"
    FLAG_MONTO_CERO      = "flag_monto_cero"
    FLAG_PRECIO_EXT      = "flag_precio_extremo"
    FLAG_CLI_ESPECIAL    = "flag_cliente_especial"
    FLAG_DOS_SIN_LEC     = "flag_dos_sin_lectura"
    FLAG_TRES_SIN_LEC    = "flag_tres_o_mas_sin_lectura"
    FLAG_ULT_SIN_LEC     = "flag_ultimo_sin_lectura"
    FLAG_EXCEDE_POT      = "flag_consumo_excede_potencia"
    FLAG_CERCA_POT       = "flag_consumo_cercano_potencia"
    FLAG_POT_FALTANTE    = "flag_potencia_faltante"
    FLAG_PER_IRREG       = "flag_periodo_irregular"
    FLAG_PER_MENS_IRREG  = "flag_periodo_mensual_irregular"
    FLAG_PER_BIM_IRREG   = "flag_periodo_bimestral_irregular"
    FLAG_AUM_30          = "flag_aumento_consumo_30"
    FLAG_AUM_100         = "flag_aumento_consumo_100"
    FLAG_AUM_INTER_30    = "flag_aumento_interanual_30"
    FLAG_AUM_INTER_100   = "flag_aumento_interanual_100"
    FLAG_AUM_PROM_3M     = "flag_aumento_vs_promedio_3m_30"
    FLAG_MULT_MED        = "flag_multiples_medidores"
    FLAG_CAMBIO_MED      = "flag_cambio_medidor"
    FLAG_TASA_ALTA       = "flag_tasa_interes_alta"
    FLAG_TASA_EXT        = "flag_tasa_interes_extrema"

    # ── Columnas booleanas (no tienen prefijo flag_) ─────────────────────────────
    COL_SALDO_ANT        = "tiene_saldo_anterior"
    COL_TIENE_INTERES    = "tiene_interes"
    COL_ABONO_ELECTRODEP = "abono_electrodependiente"
    COL_ABONO_CNR        = "abono_cnr"
    COL_ABONO_REACTIVO   = "abono_consumo_reactivo"
    COL_ABONO_SUBSIDIO   = "abono_subsidio"
    COL_CONSEC_SIN_LEC   = "consecutivo_sin_lectura"

    # Fecha base de Excel (serial numérico → fecha real)
    _EXCEL_EPOCH = datetime(1899, 12, 30)

    def __init__(self, ruta_excel: str):
        """
        Parámetros
        ──────────
        ruta_excel  Ruta al archivo .xlsx con la base de boletas.
        """
        self._ruta = ruta_excel
        self._df_completo: pd.DataFrame = self._cargar_excel()
        self._df_cliente: Optional[pd.DataFrame] = None
        self._id_cliente: Optional[int] = None

        print(f"[LectorBoletas] {len(self._df_completo):,} registros cargados. "
              f"Clientes únicos: {self._df_completo[self.COL_CLIENTE].nunique():,}")

    # ── Carga ────────────────────────────────────────────────────────────────

    def _cargar_excel(self) -> pd.DataFrame:
        """
        Carga el Excel. Las columnas de fecha ya vienen como datetime64[us]
        cuando openpyxl las interpreta correctamente — no se reconvierten.
        Solo se normalizan los flags booleanos que pueden llegar como strings.
        """
        df = pd.read_excel(self._ruta)

        # Asegurar que las columnas de fecha sean datetime (por si acaso llegan
        # como object en algún entorno); si ya son datetime64 esto es un no-op.
        date_cols = [self.COL_FECHA_EMISION, self.COL_FECHA_LEC_ACT, self.COL_FECHA_LEC_ANT]
        for col in date_cols:
            if col in df.columns and not pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = pd.to_datetime(df[col], errors="coerce")

        # Normalizar flags booleanos (pueden llegar como "TRUE"/"FALSE" strings)
        flag_cols = [
            self.FLAG_SIN_LECTURA, self.FLAG_CONS_CERO, self.FLAG_CONS_NEG,
            self.FLAG_OUTLIER, self.FLAG_MONTO_CERO, self.FLAG_PRECIO_EXT,
            self.FLAG_CLI_ESPECIAL, self.FLAG_DOS_SIN_LEC, self.FLAG_TRES_SIN_LEC,
            self.FLAG_ULT_SIN_LEC, self.FLAG_EXCEDE_POT, self.FLAG_CERCA_POT,
            self.FLAG_POT_FALTANTE, self.FLAG_PER_IRREG, self.FLAG_PER_MENS_IRREG,
            self.FLAG_PER_BIM_IRREG, self.FLAG_AUM_30, self.FLAG_AUM_100,
            self.FLAG_AUM_INTER_30, self.FLAG_AUM_INTER_100, self.FLAG_AUM_PROM_3M,
            self.FLAG_MULT_MED, self.FLAG_CAMBIO_MED, self.FLAG_TASA_ALTA,
            self.FLAG_TASA_EXT, self.COL_SALDO_ANT, self.COL_TIENE_INTERES,
            self.COL_ABONO_ELECTRODEP, self.COL_ABONO_CNR, self.COL_ABONO_REACTIVO,
            self.COL_ABONO_SUBSIDIO, "abono_devolucion_provisorio",
            "abono_subsidio_pandemia", "nromed",
            "flag_aumento_consumo", "flag_disminucion_consumo",
            "flag_aumento_interanual", "flag_disminucion_interanual",
        ]
        bool_map = {"TRUE": True, "FALSE": False, True: True, False: False}
        for col in flag_cols:
            if col in df.columns:
                df[col] = df[col].map(bool_map).fillna(False)

        return df

    def cargar_cliente(self, id_cliente: int) -> "LectorBoletas":
        """
        Filtra todos los registros del cliente y los ordena por fecha de emisión.

        Parámetros
        ──────────
        id_cliente  Número de cliente (campo `cliente` del Excel).

        Retorna self para encadenamiento de métodos.
        """
        df = self._df_completo[self._df_completo[self.COL_CLIENTE] == id_cliente].copy()

        if df.empty:
            raise ValueError(f"Cliente {id_cliente} no encontrado en la base de datos.")

        df = df.sort_values(self.COL_FECHA_EMISION).reset_index(drop=True)
        self._df_cliente = df
        self._id_cliente = id_cliente

        boletas     = df[df[self.COL_TIPO_DOC] == "Boleta"]
        notas_cred  = df[df[self.COL_TIPO_DOC] == "NotaCredito"]
        print(f"[cargar_cliente] Cliente {id_cliente}: "
              f"{len(boletas)} boletas | {len(notas_cred)} notas de crédito | "
              f"Período: {df[self.COL_FECHA_EMISION].min().date()} → "
              f"{df[self.COL_FECHA_EMISION].max().date()}")
        return self

    def _verificar_cliente(self) -> pd.DataFrame:
        """Valida que haya un cliente cargado y retorna su DataFrame de boletas."""
        if self._df_cliente is None:
            raise RuntimeError("Ejecuta cargar_cliente(id) primero.")
        return self._df_cliente[self._df_cliente[self.COL_TIPO_DOC] == "Boleta"].copy()

    # ── 1. Resumen general ───────────────────────────────────────────────────

    def resumen_cliente(self) -> dict:
        """
        Retorna un dict con la información general del cliente:
        empresa, tarifa, sector, período cubierto, medidor y totales globales.
        """
        df = self._verificar_cliente()
        ultima = df.iloc[-1]

        resumen = {
            "id_cliente":         self._id_cliente,
            "empresa":            df[self.COL_EMPRESA].iloc[0],
            "tipo_tarifa":        ultima.get(self.COL_TARIFA, "N/A"),
            "sector_tarifario":   ultima.get(self.COL_SECTOR, "N/A"),
            "potencia_conectada_kva": ultima.get(self.COL_POTENCIA, None),
            "fecha_primera_boleta": df[self.COL_FECHA_EMISION].min().date(),
            "fecha_ultima_boleta":  df[self.COL_FECHA_EMISION].max().date(),
            "total_boletas":      len(df),
            "notas_credito":      len(self._df_cliente[
                                        self._df_cliente[self.COL_TIPO_DOC] == "NotaCredito"]),
            "consumo_total_kwh":  round(df[self.COL_CONSUMO].sum(), 2),
            "monto_total_facturado_clp": round(df[self.COL_MONTO_TOTAL].sum(), 0),
            "tiene_arriendo_medidor": bool(df[self.COL_MONTO_ARRIENDO].sum() > 0),
            "es_cliente_especial":    bool(df[self.FLAG_CLI_ESPECIAL].any()),
        }
        return resumen

    # ── 2. Métricas de consumo ───────────────────────────────────────────────

    def metricas_consumo(self) -> dict:
        """
        Calcula estadísticas descriptivas del consumo mensual:
        promedio, mediana, desviación estándar, mín/máx y tendencia reciente.
        """
        df = self._verificar_cliente()
        consumos = df[self.COL_CONSUMO]
        consumo_diario = df[self.COL_CONSUMO_DIARIO]

        # Comparación consumo real vs proyectado
        df_con_proy = df[df[self.COL_CONSUMO_PROY] > 0].copy()
        if not df_con_proy.empty:
            df_con_proy["desv_vs_proyectado"] = (
                df_con_proy[self.COL_CONSUMO_DIARIO] - df_con_proy[self.COL_CONSUMO_PROY]
            )
            desv_proy_media = round(df_con_proy["desv_vs_proyectado"].mean(), 3)
            desv_proy_max   = round(df_con_proy["desv_vs_proyectado"].abs().max(), 3)
        else:
            desv_proy_media = desv_proy_max = None

        # Tendencia: comparar último tercio vs primer tercio de boletas
        n = len(df)
        tercio = max(1, n // 3)
        tend_inicio = consumos.iloc[:tercio].mean()
        tend_fin    = consumos.iloc[-tercio:].mean()

        metricas = {
            "consumo_promedio_kwh":      round(consumos.mean(), 2),
            "consumo_mediana_kwh":       round(consumos.median(), 2),
            "consumo_std_kwh":           round(consumos.std(), 2),
            "consumo_minimo_kwh":        round(consumos.min(), 2),
            "consumo_maximo_kwh":        round(consumos.max(), 2),
            "consumo_diario_promedio":   round(consumo_diario.mean(), 3),
            "consumo_diario_std":        round(consumo_diario.std(), 3),
            "consumo_teorico_maximo":    round(df[self.COL_CONSUMO_TEOR].iloc[-1], 2),
            "desv_vs_proyectado_media":  desv_proy_media,
            "desv_vs_proyectado_max":    desv_proy_max,
            "tendencia_consumo":         "creciente" if tend_fin > tend_inicio * 1.1
                                         else "decreciente" if tend_fin < tend_inicio * 0.9
                                         else "estable",
            "coef_variacion_pct":        round((consumos.std() / consumos.mean() * 100)
                                               if consumos.mean() > 0 else 0, 1),
            "boletas_consumo_cero":      int(df[self.FLAG_CONS_CERO].sum()),
            "boletas_consumo_negativo":  int(df[self.FLAG_CONS_NEG].sum()),
            "boletas_outlier":           int(df[self.FLAG_OUTLIER].sum()),
        }
        return metricas

    # ── 3. Discrepancias de medición ─────────────────────────────────────────

    def discrepancias_medicion(self) -> dict:
        """
        Detecta inconsistencias entre los días facturados y los días de lectura,
        boletas emitidas sin lectura real y saltos anómalos en el precio por kWh.

        Retorna un dict con métricas y una lista de boletas problemáticas.
        """
        df = self._verificar_cliente()

        # A. Diferencia entre días facturados y días entre lecturas
        df = df.copy()
        df["dif_dias"] = (df[self.COL_DIAS_FACT] - df[self.COL_DIAS_LECTURA]).abs()
        umbral_dias    = 5  # más de 5 días de diferencia → alerta
        mask_dif_dias  = df["dif_dias"] > umbral_dias

        # B. Boletas sin lectura real
        mask_sin_lec = df[self.FLAG_SIN_LECTURA].astype(bool)

        # C. Precio kWh fuera del rango esperado (3 sigmas)
        precios = df.loc[df[self.COL_PRECIO_KWH] > 0, self.COL_PRECIO_KWH]
        if len(precios) >= 3:
            p_media = precios.mean()
            p_std   = precios.std()
            mask_precio_ext = (
                (df[self.COL_PRECIO_KWH] > p_media + 3 * p_std) |
                (df[self.COL_PRECIO_KWH].between(1, p_media - 3 * p_std))
            ) & (df[self.COL_PRECIO_KWH] > 0)
        else:
            mask_precio_ext = df[self.FLAG_PRECIO_EXT].astype(bool)

        # D. Períodos de lectura anómalos (>60 días o <15 días)
        mask_per_largo = df[self.COL_DIAS_LECTURA] > 60
        mask_per_corto = (df[self.COL_DIAS_LECTURA] < 15) & (df[self.COL_DIAS_LECTURA] > 0)

        boletas_problematicas = df[
            mask_dif_dias | mask_sin_lec | mask_precio_ext | mask_per_largo | mask_per_corto
        ][[
            self.COL_NUMDOC, self.COL_FECHA_EMISION, self.COL_CONSUMO,
            self.COL_DIAS_FACT, self.COL_DIAS_LECTURA, "dif_dias",
            self.COL_PRECIO_KWH, self.FLAG_SIN_LECTURA,
        ]].copy()

        boletas_problematicas["alerta"] = boletas_problematicas.apply(
            lambda r: ", ".join(filter(None, [
                f"Δdías={int(r['dif_dias'])}"    if r["dif_dias"] > umbral_dias else "",
                "sin_lectura"                    if r[self.FLAG_SIN_LECTURA]   else "",
                "precio_extremo"                 if mask_precio_ext[r.name]    else "",
                "período_largo"                  if r[self.COL_DIAS_LECTURA] > 60 else "",
                "período_corto"                  if 0 < r[self.COL_DIAS_LECTURA] < 15 else "",
            ])), axis=1
        )

        discrepancias = {
            "boletas_con_diferencia_dias":        int(mask_dif_dias.sum()),
            "boletas_sin_lectura_real":           int(mask_sin_lec.sum()),
            "boletas_precio_extremo":             int(mask_precio_ext.sum()),
            "boletas_periodo_largo_gt60d":        int(mask_per_largo.sum()),
            "boletas_periodo_corto_lt15d":        int(mask_per_corto.sum()),
            "diferencia_dias_promedio":           round(df["dif_dias"].mean(), 1),
            "precio_kwh_promedio":                round(precios.mean(), 2) if len(precios) else None,
            "precio_kwh_std":                     round(precios.std(), 2)  if len(precios) else None,
            "total_boletas_problematicas":        len(boletas_problematicas),
            "detalle_boletas_problematicas":      boletas_problematicas.to_dict("records"),
        }
        return discrepancias

    # ── 4. Historial de facturación ──────────────────────────────────────────

    def historial_facturacion(self, ultimas_n: int = 12) -> pd.DataFrame:
        """
        Retorna un DataFrame con el historial de boletas ordenado por fecha,
        con las columnas más relevantes para análisis de reclamos.

        Parámetros
        ──────────
        ultimas_n  Número de boletas más recientes a retornar (default 12).
        """
        df = self._verificar_cliente()

        cols_salida = [
            self.COL_NUMDOC, self.COL_FECHA_EMISION,
            self.COL_CONSUMO, self.COL_CONSUMO_DIARIO,
            self.COL_DIAS_FACT, self.COL_DIAS_LECTURA,
            self.COL_MONTO_ELEC, self.COL_MONTO_TOTAL,
            self.COL_PRECIO_KWH,
            self.COL_MONTO_ARRIENDO, self.COL_MONTO_DEVOL, self.COL_MONTO_COMP,
            self.COL_MONTO_SUBS, self.COL_MONTO_PAND, self.COL_MONTO_CUOTA,
            self.FLAG_SIN_LECTURA, self.FLAG_CONS_CERO,
            self.FLAG_OUTLIER, self.FLAG_PRECIO_EXT,
        ]
        cols_salida = [c for c in cols_salida if c in df.columns]

        return df[cols_salida].tail(ultimas_n).reset_index(drop=True)

    # ── 5. Detección de anomalías ────────────────────────────────────────────

    def detectar_anomalias(self) -> dict:
        """
        Consolida todos los flags de anomalía del cliente y calcula
        métricas sobre devoluciones, compensaciones y subsidios recibidos.
        """
        df = self._verificar_cliente()
        df_all = self._df_cliente  # incluye notas de crédito

        # Montos de ajustes y devoluciones
        total_devol  = df[self.COL_MONTO_DEVOL].sum()
        total_comp   = df[self.COL_MONTO_COMP].sum()
        total_subs   = df[self.COL_MONTO_SUBS].sum()
        total_arr    = df[self.COL_MONTO_ARRIENDO].sum()
        total_nc     = df_all.loc[
            df_all[self.COL_TIPO_DOC] == "NotaCredito", self.COL_MONTO_TOTAL
        ].sum()

        # Boleta con mayor cobro (candidata al reclamo)
        idx_max  = df[self.COL_MONTO_TOTAL].idxmax()
        boleta_max = df.loc[idx_max]

        # Saltos bruscos en monto (>150% del promedio histórico)
        promedio_monto = df[self.COL_MONTO_TOTAL].mean()
        mask_salto     = df[self.COL_MONTO_TOTAL] > promedio_monto * 1.5
        boletas_salto  = df[mask_salto][[
            self.COL_NUMDOC, self.COL_FECHA_EMISION,
            self.COL_MONTO_TOTAL, self.COL_CONSUMO
        ]].to_dict("records")

        anomalias = {
            # Flags acumulados
            "total_flags_sin_lectura":    int(df[self.FLAG_SIN_LECTURA].sum()),
            "total_flags_consumo_cero":   int(df[self.FLAG_CONS_CERO].sum()),
            "total_flags_consumo_neg":    int(df[self.FLAG_CONS_NEG].sum()),
            "total_flags_outlier":        int(df[self.FLAG_OUTLIER].sum()),
            "total_flags_monto_cero":     int(df[self.FLAG_MONTO_CERO].sum()),
            "total_flags_precio_extremo": int(df[self.FLAG_PRECIO_EXT].sum()),
            # Devoluciones y ajustes
            "monto_total_devoluciones_clp":     round(total_devol, 0),
            "monto_total_compensaciones_clp":   round(total_comp, 0),
            "monto_total_subsidios_clp":        round(total_subs, 0),
            "monto_total_arriendo_medidor_clp": round(total_arr, 0),
            "monto_total_notas_credito_clp":    round(total_nc, 0),
            # Boleta más alta
            "boleta_monto_maximo": {
                "numdocumento": boleta_max[self.COL_NUMDOC],
                "fecha":        str(boleta_max[self.COL_FECHA_EMISION].date()),
                "monto_clp":    round(boleta_max[self.COL_MONTO_TOTAL], 0),
                "consumo_kwh":  round(boleta_max[self.COL_CONSUMO], 2),
            },
            # Boletas con saltos bruscos
            "boletas_salto_monto_gt150pct_promedio": boletas_salto,
        }
        return anomalias

    # ── 6. Contexto para el agente clasificador ──────────────────────────────

    def generar_contexto_reclamo(self) -> dict:
        """
        Consolida toda la información en un único dict listo para alimentar
        al AgenteClasificadorSEC, con todas las alertas posibles activadas.
        """
        resumen   = self.resumen_cliente()
        consumo   = self.metricas_consumo()
        discrepan = self.discrepancias_medicion()
        anomalias = self.detectar_anomalias()
        historial = self.historial_facturacion(ultimas_n=6).to_dict("records")

        df = self._verificar_cliente()  # solo boletas

        def cualquiera(flag):
            """True si al menos una boleta tiene ese flag activo."""
            return bool(df[flag].any()) if flag in df.columns else False

        def maximo(col):
            return df[col].max() if col in df.columns else 0

        alertas = []

        # ── Lectura y medición ───────────────────────────────────────────────────
        if cualquiera(self.FLAG_SIN_LECTURA):
            alertas.append("boletas_sin_lectura_real")
        if cualquiera(self.FLAG_DOS_SIN_LEC):
            alertas.append("dos_periodos_consecutivos_sin_lectura")
        if cualquiera(self.FLAG_TRES_SIN_LEC):
            alertas.append("tres_o_mas_periodos_sin_lectura")
        if cualquiera(self.FLAG_ULT_SIN_LEC):
            alertas.append("ultima_boleta_sin_lectura")
        if maximo(self.COL_CONSEC_SIN_LEC) >= 2:
            alertas.append(f"maximo_{int(maximo(self.COL_CONSEC_SIN_LEC))}_sin_lectura_seguidos")

        # ── Consumo ──────────────────────────────────────────────────────────────
        if consumo["boletas_outlier"] > 0:
            alertas.append("consumo_outlier_detectado")
        if consumo["boletas_consumo_cero"] > 0:
            alertas.append("boletas_con_consumo_cero")
        if consumo["boletas_consumo_negativo"] > 0:
            alertas.append("consumo_negativo_registrado")
        if consumo["coef_variacion_pct"] > 40:
            alertas.append("consumo_alta_variabilidad")
        if cualquiera(self.FLAG_AUM_30):
            alertas.append("aumento_consumo_mayor_30pct_mensual")
        if cualquiera(self.FLAG_AUM_100):
            alertas.append("aumento_consumo_mayor_100pct_mensual")
        if cualquiera(self.FLAG_AUM_INTER_30):
            alertas.append("aumento_consumo_mayor_30pct_interanual")
        if cualquiera(self.FLAG_AUM_INTER_100):
            alertas.append("aumento_consumo_mayor_100pct_interanual")
        if cualquiera(self.FLAG_AUM_PROM_3M):
            alertas.append("aumento_consumo_mayor_30pct_vs_promedio_3m")

        # ── Potencia ─────────────────────────────────────────────────────────────
        if cualquiera(self.FLAG_EXCEDE_POT):
            alertas.append("consumo_excede_potencia_contratada")
        if cualquiera(self.FLAG_CERCA_POT):
            alertas.append("consumo_cercano_limite_potencia")
        if cualquiera(self.FLAG_POT_FALTANTE):
            alertas.append("potencia_contratada_no_registrada")

        # ── Período de facturación ───────────────────────────────────────────────
        if discrepan["boletas_con_diferencia_dias"] > 0:
            alertas.append("discrepancia_dias_facturados_vs_lectura")
        if discrepan["boletas_periodo_largo_gt60d"] > 0:
            alertas.append("periodo_facturacion_largo_mayor_60_dias")
        if discrepan["boletas_periodo_corto_lt15d"] > 0:
            alertas.append("periodo_facturacion_corto_menor_15_dias")
        if cualquiera(self.FLAG_PER_IRREG):
            alertas.append("periodo_de_facturacion_irregular")
        if cualquiera(self.FLAG_PER_MENS_IRREG):
            alertas.append("irregularidad_en_ciclo_mensual")
        if cualquiera(self.FLAG_PER_BIM_IRREG):
            alertas.append("irregularidad_en_ciclo_bimestral")

        # ── Precio ───────────────────────────────────────────────────────────────
        if discrepan["boletas_precio_extremo"] > 0:
            alertas.append("precio_kwh_anomalo_detectado")

        # ── Medidor ──────────────────────────────────────────────────────────────
        if cualquiera(self.FLAG_CAMBIO_MED):
            alertas.append("cambio_de_medidor_registrado")
        if cualquiera(self.FLAG_MULT_MED):
            n = int(df["num_medidores_distintos"].max()) if "num_medidores_distintos" in df.columns else "?"
            alertas.append(f"multiples_medidores_detectados_{n}")

        # ── Financiero ───────────────────────────────────────────────────────────
        if anomalias["monto_total_devoluciones_clp"] > 0:
            alertas.append("devoluciones_provisorias_registradas")
        if anomalias["monto_total_notas_credito_clp"] > 0:
            alertas.append("notas_de_credito_emitidas")
        if anomalias["monto_total_compensaciones_clp"] > 0:
            alertas.append("compensaciones_aplicadas")
        if cualquiera(self.COL_SALDO_ANT):
            alertas.append("saldo_anterior_pendiente_en_boletas")
        if cualquiera(self.COL_TIENE_INTERES):
            alertas.append("intereses_por_mora_aplicados")
        if cualquiera(self.FLAG_TASA_ALTA):
            alertas.append("tasa_interes_mora_alta")
        if cualquiera(self.FLAG_TASA_EXT):
            alertas.append("tasa_interes_mora_extrema")
        if anomalias["monto_total_subsidios_clp"] > 0:
            alertas.append("subsidios_registrados")

        # ── Abonos especiales ────────────────────────────────────────────────────
        if cualquiera(self.COL_ABONO_ELECTRODEP):
            alertas.append("abono_cliente_electrodependiente")
        if cualquiera(self.COL_ABONO_CNR):
            alertas.append("abono_cnr_aplicado")
        if cualquiera(self.COL_ABONO_REACTIVO):
            alertas.append("cobro_consumo_reactivo")
        if cualquiera("abono_subsidio_pandemia"):
            alertas.append("subsidio_pandemia_registrado")
        if cualquiera("abono_cuota_pandemia"):
            alertas.append("cuota_pandemia_registrada")

        contexto = {
            "cliente":      resumen,
            "consumo":      consumo,
            "discrepancias_medicion": {
                k: v for k, v in discrepan.items()
                if k != "detalle_boletas_problematicas"
            },
            "anomalias": {
                k: v for k, v in anomalias.items()
                if k != "boletas_salto_monto_gt150pct_promedio"
            },
            "alertas_activas": alertas,
            "ultimas_6_boletas": historial,
        }
        return contexto

    # ── Representación ───────────────────────────────────────────────────────

    def __repr__(self) -> str:
        estado = f"cliente={self._id_cliente}" if self._id_cliente else "sin cliente cargado"
        return f"LectorBoletas(archivo='{self._ruta}', {estado})"
    
    def imprimir_contexto(self) -> None:
        ctx = self.generar_contexto_reclamo()

        def seccion(titulo):
            print(f"\n{'=' * 55}")
            print(f"  {titulo}")
            print(f"{'=' * 55}")

        def fila(clave, valor):
            print(f"  {clave:<32} {valor}")

        def separador():
            print(f"  {'-' * 50}")

        # ── 1. CLIENTE ──────────────────────────────────────────
        c = ctx["cliente"]
        seccion("① CLIENTE")
        fila("ID Cliente:",              c["id_cliente"])
        fila("Empresa:",                 c["empresa"])
        fila("Tipo de tarifa:",          c["tipo_tarifa"])
        fila("Sector tarifario:",        c["sector_tarifario"])
        fila("Potencia contratada:",     f"{c['potencia_conectada_kva']} kVA")
        fila("Período cubierto:",        f"{c['fecha_primera_boleta']}  →  {c['fecha_ultima_boleta']}")
        fila("Boletas / Notas crédito:", f"{c['total_boletas']} boletas  /  {c['notas_credito']} NC")
        fila("Consumo total:",           f"{c['consumo_total_kwh']:,.2f} kWh")
        fila("Monto total facturado:",   f"${c['monto_total_facturado_clp']:,.0f}")
        fila("Arriendo medidor:",        "Sí" if c["tiene_arriendo_medidor"] else "No")
        fila("Cliente especial:",        "Sí" if c["es_cliente_especial"] else "No")

        # ── 2. CONSUMO ──────────────────────────────────────────
        m = ctx["consumo"]
        seccion("② MÉTRICAS DE CONSUMO")
        fila("Promedio mensual:",        f"{m['consumo_promedio_kwh']:,.2f} kWh")
        fila("Mediana:",                 f"{m['consumo_mediana_kwh']:,.2f} kWh")
        fila("Desv. estándar:",          f"{m['consumo_std_kwh']:,.2f} kWh")
        fila("Mín / Máx:",              f"{m['consumo_minimo_kwh']:,.2f}  /  {m['consumo_maximo_kwh']:,.2f} kWh")
        fila("Coef. de variación:",      f"{m['coef_variacion_pct']} %")
        fila("Consumo diario promedio:", f"{m['consumo_diario_promedio']} kWh/día")
        fila("Consumo teórico máximo:",  f"{m['consumo_teorico_maximo']:,.2f} kWh")
        fila("Tendencia:",               m["tendencia_consumo"])
        separador()
        fila("Boletas consumo cero:",    m["boletas_consumo_cero"])
        fila("Boletas consumo negativo:",m["boletas_consumo_negativo"])
        fila("Boletas outlier:",         m["boletas_outlier"])

        # ── 3. DISCREPANCIAS ────────────────────────────────────
        d = ctx["discrepancias_medicion"]
        seccion("③ DISCREPANCIAS DE MEDICIÓN")
        fila("Sin lectura real:",           d["boletas_sin_lectura_real"])
        fila("Diferencia días factura:",    d["boletas_con_diferencia_dias"])
        fila("Precio kWh anómalo:",         d["boletas_precio_extremo"])
        fila("Período largo (> 60 días):",  d["boletas_periodo_largo_gt60d"])
        fila("Período corto (< 15 días):",  d["boletas_periodo_corto_lt15d"])
        separador()
        fila("Diferencia días promedio:",   f"{d['diferencia_dias_promedio']} días")
        fila("Precio kWh promedio:",        f"${d['precio_kwh_promedio']:,.2f}" if d["precio_kwh_promedio"] else "N/A")
        fila("Total boletas irregulares:",  d["total_boletas_problematicas"])

        # ── 4. ANOMALÍAS ────────────────────────────────────────
        a = ctx["anomalias"]
        bm = a["boleta_monto_maximo"]
        seccion("④ ANOMALÍAS Y AJUSTES")
        fila("Boleta monto máximo:",     f"N°{bm['numdocumento']}  |  {bm['fecha']}  |  ${bm['monto_clp']:,.0f}  ({bm['consumo_kwh']:,.2f} kWh)")
        separador()
        fila("Total devoluciones:",      f"${a['monto_total_devoluciones_clp']:,.0f}")
        fila("Total compensaciones:",    f"${a['monto_total_compensaciones_clp']:,.0f}")
        fila("Total subsidios:",         f"${a['monto_total_subsidios_clp']:,.0f}")
        fila("Total notas de crédito:",  f"${a['monto_total_notas_credito_clp']:,.0f}")
        fila("Arriendo medidor total:",  f"${a['monto_total_arriendo_medidor_clp']:,.0f}")

        # ── 5. ALERTAS ──────────────────────────────────────────
        alertas = ctx["alertas_activas"]
        seccion("⑤ ALERTAS ACTIVAS")
        if alertas:
            for alerta in alertas:
                print(f"  - {alerta.replace('_', ' ')}")
        else:
            print("  Sin alertas activas.")

        # ── 6. ÚLTIMAS BOLETAS ──────────────────────────────────
        seccion("⑥ ÚLTIMAS 6 BOLETAS")
        print(f"  {'Fecha':<13} {'Doc N°':<13} {'kWh':>8} {'Monto':>11} {'$/kWh':>7} {'Días':>5}  Flags")
        separador()
        for b in ctx["ultimas_6_boletas"]:
            fecha   = str(b.get("fechaemision", ""))[:10]
            ndoc    = str(b.get("numdocumento", ""))
            consumo = f"{b.get('consumo_total', 0):>8.1f}"
            monto   = f"${int(b.get('montototal_doc', 0)):>9,}"
            precio  = f"{b.get('precio_kwh', 0):>7.1f}"
            dias    = f"{int(b.get('dias_facturacion', 0)):>5}"
            flags   = "  ".join(filter(None, [
                "SIN_LEC"  if b.get("flag_sin_lectura")      else "",
                "CONS_0"   if b.get("flag_consumo_cero")     else "",
                "OUTLIER"  if b.get("flag_consumo_outlier")  else "",
                "PRECIO!"  if b.get("flag_precio_extremo")   else "",
            ])) or "OK"
            print(f"  {fecha:<13} {ndoc:<13} {consumo} {monto:>11} {precio} {dias}  {flags}")

        print(f"\n{'=' * 55}\n")   
    
# Instanciar y cargar cliente
lector = LectorBoletas("base_recl_LectorBLT_IA.xlsx")
lector.cargar_cliente(1236999)

# Ver historial tabular de las últimas 12 boletas
df_hist = lector.historial_facturacion(ultimas_n=12)

# Generar contexto consolidado para el agente
lector.imprimir_contexto()