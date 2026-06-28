import sys
import io
import contextlib
import traceback
from datetime import datetime

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QTextCursor, QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QProgressBar, QPlainTextEdit, QFrame,
    QGroupBox, QSizePolicy, QScrollArea, QMessageBox, QSpacerItem
)
from Extractor     import extraer,   ID_REFERENCIA as ID_REFERENCIA_DEFAULT, ID_PENDIENTES
from Clasificador  import clasificar
from LectorBoletas import analizar,  EXCEL_PATH, TIPOLOGIA_NOMBRE

ID_INCIDENTE_DEFAULT = "2835115"

# ── PALETA Y ESTILOS ──────────────────────────────────────────────────────────────

COLOR_BG          = "#0f1117"
COLOR_PANEL       = "#161922"
COLOR_PANEL_ALT   = "#1b1f2a"
COLOR_BORDER      = "#2a2f3d"
COLOR_TEXT        = "#e6e9ef"
COLOR_TEXT_DIM    = "#8b93a7"
COLOR_ACCENT      = "#3b82f6"
COLOR_ACCENT_DIM  = "#1e3a5f"
COLOR_GREEN       = "#22c55e"
COLOR_RED         = "#ef4444"
COLOR_YELLOW      = "#eab308"
COLOR_ORANGE      = "#f97316"
COLOR_GREY        = "#6b7280"

ESTADO_COLOR = {
    "VALIDADO_PARA_DESPACHO":    COLOR_GREEN,
    "CLASIFICADO_PARA_REVISION": COLOR_YELLOW,
    "NO_EVALUABLE":              COLOR_GREY,
    "BLOQUEADO_POR_CONTROL":     COLOR_RED,
    "SIN_SENAL_COMPATIBLE":      COLOR_ORANGE,
}

MONO_FONT = "Consolas, 'Cascadia Mono', 'Courier New', monospace"


# ── HILO DE TRABAJO ───────────────────────────────────────────────────────────────

class PipelineWorker(QThread):
    log_emitted      = pyqtSignal(str)
    stage_started    = pyqtSignal(int, str)     # (índice de etapa, nombre)
    stage_finished   = pyqtSignal(int)          # índice de etapa completada
    finished_ok      = pyqtSignal(dict)         # resultado final completo
    finished_error    = pyqtSignal(str, str)     # (etapa, mensaje de error)

    def __init__(self, n_referencia: str, id_incidente: str, id_pendientes: int, excel_path: str):
        super().__init__()
        self.n_referencia  = n_referencia
        self.id_incidente  = id_incidente
        self.id_pendientes = id_pendientes
        self.excel_path    = excel_path

    def run(self):
        try:
            # ── ETAPA 1 — EXTRACTOR ──────────────────────────────────────────────
            self.stage_started.emit(0, "Extractor SEC")
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    datos_caso = extraer(
                        ID_REFERENCIA=self.n_referencia,
                        id_pendientes=self.id_pendientes,
                    )
            except Exception as e:
                self.log_emitted.emit(buf.getvalue())
                raise RuntimeError(f"[Extractor] {e}") from e
            self.log_emitted.emit(buf.getvalue())
            self.stage_finished.emit(0)

            if self.id_incidente:
                datos_caso = dict(datos_caso)
                datos_caso["_id_incidente_excel"] = self.id_incidente

            # ── ETAPA 2 — CLASIFICADOR ───────────────────────────────────────────
            self.stage_started.emit(1, "Clasificador IA")
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    clasificacion = clasificar(datos_caso)
            except Exception as e:
                self.log_emitted.emit(buf.getvalue())
                raise RuntimeError(f"[Clasificador] {e}") from e
            self.log_emitted.emit(buf.getvalue())
            self.stage_finished.emit(1)

            # ── ETAPA 3 — LECTOR DE BOLETAS ──────────────────────────────────────
            self.stage_started.emit(2, "Lector de Boletas")

            datos_para_lector = dict(datos_caso)
            if self.id_incidente:
                datos_para_lector["id"] = self.id_incidente

            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    reporte = analizar(
                        datos_para_lector, clasificacion,
                        excel_path=self.excel_path,
                    )
            except Exception as e:
                self.log_emitted.emit(buf.getvalue())
                raise RuntimeError(f"[LectorBoletas] {e}") from e
            self.log_emitted.emit(buf.getvalue())
            self.stage_finished.emit(2)

            self.finished_ok.emit({
                "caso":          datos_caso,
                "clasificacion": clasificacion,
                "lector":        reporte,
            })

        except RuntimeError as e:
            etapa = str(e).split("]")[0].strip("[")
            self.finished_error.emit(etapa, str(e))
        except Exception as e:
            self.finished_error.emit("Desconocida", f"{e}\n{traceback.format_exc()}")

class StageCard(QFrame):
    PENDIENTE  = 0
    EN_CURSO   = 1
    COMPLETADA = 2
    ERROR      = 3

    def __init__(self, numero: int, titulo: str, subtitulo: str, parent=None):
        super().__init__(parent)
        self.numero = numero
        self.estado = self.PENDIENTE

        self.setFrameShape(QFrame.NoFrame)
        self.setMinimumHeight(72)
        self.setObjectName("stageCard")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(12)

        self.badge = QLabel(str(numero))
        self.badge.setFixedSize(34, 34)
        self.badge.setAlignment(Qt.AlignCenter)
        self.badge.setObjectName("stageBadge")

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        self.lbl_titulo = QLabel(titulo)
        self.lbl_titulo.setObjectName("stageTitulo")
        self.lbl_subtitulo = QLabel(subtitulo)
        self.lbl_subtitulo.setObjectName("stageSubtitulo")
        text_box.addWidget(self.lbl_titulo)
        text_box.addWidget(self.lbl_subtitulo)

        self.lbl_estado = QLabel("Pendiente")
        self.lbl_estado.setObjectName("stageEstado")
        self.lbl_estado.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        layout.addWidget(self.badge)
        layout.addLayout(text_box, 1)
        layout.addWidget(self.lbl_estado)

        self._aplicar_estado()

    def set_estado(self, estado: int):
        self.estado = estado
        self._aplicar_estado()

    def _aplicar_estado(self):
        if self.estado == self.PENDIENTE:
            color_badge, color_text, texto = COLOR_BORDER, COLOR_TEXT_DIM, "Pendiente"
        elif self.estado == self.EN_CURSO:
            color_badge, color_text, texto = COLOR_ACCENT, COLOR_ACCENT, "En curso…"
        elif self.estado == self.COMPLETADA:
            color_badge, color_text, texto = COLOR_GREEN, COLOR_GREEN, "Completado ✓"
        else:
            color_badge, color_text, texto = COLOR_RED, COLOR_RED, "Error ✗"

        self.badge.setStyleSheet(
            f"background-color: {color_badge}; color: #0f1117; "
            f"border-radius: 17px; font-weight: 700; font-size: 14px;"
        )
        self.lbl_estado.setStyleSheet(f"color: {color_text}; font-weight: 600; font-size: 12px;")
        self.lbl_estado.setText(texto)

class EstadoBadge(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(46)
        self.set_estado(None)

    def set_estado(self, estado):
        if estado is None:
            self.setText("— SIN EJECUTAR —")
            color = COLOR_GREY
        else:
            self.setText(estado.replace("_", " "))
            color = ESTADO_COLOR.get(estado, COLOR_GREY)
        self.setStyleSheet(
            f"background-color: {color}22; color: {color}; "
            f"border: 1.5px solid {color}; border-radius: 8px; "
            f"font-weight: 700; font-size: 15px; letter-spacing: 0.5px;"
        )


# ── VENTANA PRINCIPAL ─────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.setWindowTitle("SEC · Pipeline de Clasificación y Prerresolución de Reclamos")
        self.resize(1180, 780)
        self.setMinimumSize(980, 680)

        self._build_ui()
        self._apply_stylesheet()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(14)

        # Encabezado
        header = QVBoxLayout()
        header.setSpacing(2)
        titulo = QLabel("Sistema de Clasificación y Prerresolución de Reclamos Eléctricos")
        titulo.setObjectName("appTitulo")
        subtitulo = QLabel("SEC · Extractor → Clasificador IA → Lector de Boletas  ·  v1.0")
        subtitulo.setObjectName("appSubtitulo")
        header.addWidget(titulo)
        header.addWidget(subtitulo)
        root.addLayout(header)

        # Panel de entrada
        input_group = QGroupBox("Parámetros de ejecución")
        input_layout = QGridLayout(input_group)
        input_layout.setHorizontalSpacing(16)
        input_layout.setVerticalSpacing(10)
        input_layout.setContentsMargins(16, 16, 16, 16)

        lbl_ref = QLabel("N° de referencia SEC:")
        self.input_referencia = QLineEdit()
        self.input_referencia.setPlaceholderText("Ej: 260128-000718")
        self.input_referencia.setText(str(ID_REFERENCIA_DEFAULT))

        lbl_id = QLabel("ID de incidente (Excel boletas):")
        self.input_id_incidente = QLineEdit()
        self.input_id_incidente.setPlaceholderText("Ej: 2835115")
        self.input_id_incidente.setText(ID_INCIDENTE_DEFAULT)

        self.btn_ejecutar = QPushButton("▶  EJECUTAR PIPELINE")
        self.btn_ejecutar.setObjectName("btnEjecutar")
        self.btn_ejecutar.setMinimumHeight(40)
        self.btn_ejecutar.clicked.connect(self._on_ejecutar)

        self.btn_limpiar = QPushButton("Limpiar consola")
        self.btn_limpiar.setObjectName("btnSecundario")
        self.btn_limpiar.setMinimumHeight(40)
        self.btn_limpiar.clicked.connect(self._limpiar_consola)

        input_layout.addWidget(lbl_ref, 0, 0)
        input_layout.addWidget(self.input_referencia, 0, 1)
        input_layout.addWidget(lbl_id, 0, 2)
        input_layout.addWidget(self.input_id_incidente, 0, 3)
        input_layout.addWidget(self.btn_ejecutar, 0, 4)
        input_layout.addWidget(self.btn_limpiar,  0, 5)
        input_layout.setColumnStretch(1, 2)
        input_layout.setColumnStretch(3, 2)

        root.addWidget(input_group)

        body = QHBoxLayout()
        body.setSpacing(14)

        left_col = QVBoxLayout()
        left_col.setSpacing(12)

        stages_group = QGroupBox("Progreso del pipeline")
        stages_layout = QVBoxLayout(stages_group)
        stages_layout.setSpacing(8)
        stages_layout.setContentsMargins(14, 14, 14, 14)

        self.stage_cards = [
            StageCard(1, "Extractor SEC", "Descarga y localiza el caso en el reporte de pendientes"),
            StageCard(2, "Clasificador IA", "Clasifica el texto del reclamo vía API"),
            StageCard(3, "Lector de Boletas", "Evalúa perfiles C1-C14 y controles previos"),
        ]
        for card in self.stage_cards:
            stages_layout.addWidget(card)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 3)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Etapa %v / %m")
        stages_layout.addWidget(self.progress_bar)

        left_col.addWidget(stages_group)

        result_group = QGroupBox("Resultado final")
        result_layout = QVBoxLayout(result_group)
        result_layout.setSpacing(8)
        result_layout.setContentsMargins(14, 14, 14, 14)

        self.estado_badge = EstadoBadge()
        result_layout.addWidget(self.estado_badge)

        self.lbl_resumen = QLabel("Ejecute el pipeline para ver el resultado aquí.")
        self.lbl_resumen.setObjectName("resumenTexto")
        self.lbl_resumen.setWordWrap(True)
        self.lbl_resumen.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(self.lbl_resumen)
        result_layout.addWidget(scroll, 1)

        left_col.addWidget(result_group, 1)

        body.addLayout(left_col, 4)

        console_group = QGroupBox("Consola de ejecución")
        console_layout = QVBoxLayout(console_group)
        console_layout.setContentsMargins(10, 10, 10, 10)

        self.console = QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        self.console.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        font = QFont("Consolas")
        font.setStyleHint(QFont.Monospace)
        font.setPointSize(10)
        self.console.setFont(font)
        console_layout.addWidget(self.console)

        body.addWidget(console_group, 6)

        root.addLayout(body, 1)

        self.statusBar().showMessage("Listo. Ingrese los parámetros y presione Ejecutar.")
        self._log_sistema("Sistema iniciado. Esperando ejecución.")

    # ESTILOS

    def _apply_stylesheet(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background-color: {COLOR_BG};
                color: {COLOR_TEXT};
                font-family: Segoe UI, Arial, sans-serif;
            }}
            #appTitulo {{
                font-size: 19px;
                font-weight: 700;
                color: {COLOR_TEXT};
            }}
            #appSubtitulo {{
                font-size: 11.5px;
                color: {COLOR_TEXT_DIM};
                letter-spacing: 0.3px;
            }}
            QGroupBox {{
                background-color: {COLOR_PANEL};
                border: 1px solid {COLOR_BORDER};
                border-radius: 10px;
                margin-top: 10px;
                font-weight: 600;
                font-size: 12px;
                color: {COLOR_TEXT_DIM};
                padding-top: 6px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 6px;
                color: {COLOR_TEXT_DIM};
            }}
            QLineEdit {{
                background-color: {COLOR_PANEL_ALT};
                border: 1px solid {COLOR_BORDER};
                border-radius: 6px;
                padding: 8px 10px;
                color: {COLOR_TEXT};
                font-size: 13px;
            }}
            QLineEdit:focus {{
                border: 1px solid {COLOR_ACCENT};
            }}
            QLabel {{
                color: {COLOR_TEXT};
                font-size: 12.5px;
            }}
            #resumenTexto {{
                color: {COLOR_TEXT};
                font-size: 12.5px;
                line-height: 150%;
            }}
            QPushButton#btnEjecutar {{
                background-color: {COLOR_ACCENT};
                color: white;
                border: none;
                border-radius: 6px;
                font-weight: 700;
                font-size: 13px;
                padding: 0 18px;
            }}
            QPushButton#btnEjecutar:hover {{
                background-color: #2563eb;
            }}
            QPushButton#btnEjecutar:disabled {{
                background-color: {COLOR_GREY};
                color: #d1d5db;
            }}
            QPushButton#btnSecundario {{
                background-color: {COLOR_PANEL_ALT};
                color: {COLOR_TEXT_DIM};
                border: 1px solid {COLOR_BORDER};
                border-radius: 6px;
                font-size: 12px;
                padding: 0 14px;
            }}
            QPushButton#btnSecundario:hover {{
                background-color: {COLOR_BORDER};
                color: {COLOR_TEXT};
            }}
            QFrame#stageCard {{
                background-color: {COLOR_PANEL_ALT};
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
            }}
            #stageTitulo {{
                font-weight: 700;
                font-size: 13px;
                color: {COLOR_TEXT};
            }}
            #stageSubtitulo {{
                font-size: 10.5px;
                color: {COLOR_TEXT_DIM};
            }}
            QProgressBar {{
                background-color: {COLOR_PANEL_ALT};
                border: 1px solid {COLOR_BORDER};
                border-radius: 6px;
                text-align: center;
                color: {COLOR_TEXT};
                font-size: 11px;
                height: 22px;
            }}
            QProgressBar::chunk {{
                background-color: {COLOR_ACCENT};
                border-radius: 5px;
            }}
            QPlainTextEdit#console {{
                background-color: #0a0c10;
                color: #b8f5c0;
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                padding: 10px;
                selection-background-color: {COLOR_ACCENT_DIM};
            }}
            QScrollArea {{
                background-color: transparent;
                border: none;
            }}
            QStatusBar {{
                background-color: {COLOR_PANEL};
                color: {COLOR_TEXT_DIM};
                font-size: 11px;
            }}
            QScrollBar:vertical {{
                background-color: {COLOR_PANEL_ALT};
                width: 10px;
                border-radius: 5px;
            }}
            QScrollBar::handle:vertical {{
                background-color: {COLOR_BORDER};
                border-radius: 5px;
                min-height: 24px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: {COLOR_GREY};
            }}
        """)

    def _log_sistema(self, mensaje: str, nivel: str = "INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        color = {
            "INFO":  "#7dd3fc",
            "OK":    "#86efac",
            "WARN":  "#fde047",
            "ERROR": "#fca5a5",
        }.get(nivel, "#7dd3fc")
        self.console.appendHtml(
            f'<span style="color:#5b6472">{ts}</span>&nbsp;&nbsp;'
            f'<span style="color:{color}; font-weight:700">{nivel:<5}</span>&nbsp;'
            f'<span style="color:#e6e9ef">{self._escape(mensaje)}</span>'
        )
        self._scroll_console_to_bottom()

    def _log_pipeline(self, texto: str):
        if texto is None:
            return

        lineas = texto.splitlines()
        salida_html = []
        omitir_siguiente_detalle = False

        for linea_raw in lineas:
            linea = linea_raw.rstrip()
            stripped = linea.strip()

            if omitir_siguiente_detalle:
                omitir_siguiente_detalle = False
                if stripped.startswith("Inactivo"):
                    continue

            if stripped.startswith("○") or "[inactivo]" in stripped:
                omitir_siguiente_detalle = True

            html = self._formatear_linea_pipeline(linea)
            if html is not None:
                salida_html.append(html)

        if salida_html:
            self.console.appendHtml("<br>".join(salida_html))
        self._scroll_console_to_bottom()

    def _formatear_linea_pipeline(self, linea: str):
        stripped = linea.strip()

        if stripped == "":
            return "&nbsp;"

        n_indent = len(linea) - len(linea.lstrip(" "))
        indent_html = "&nbsp;" * min(n_indent, 6)   # tope de sangría visual
        esc = self._escape(stripped)

        # Separadores gruesos (═, ━, ==) → reducidos a una marca corta
        if all(c in "═━=" for c in stripped):
            return '<span style="color:#2a2f3d">────────────────────</span>'

        # Separadores finos (─, -, ·) → se omiten (ya hay espaciado de bloque)
        if all(c in "─\u00b7-" for c in stripped):
            return None

        # Encabezados de sección  [ TEXTO ]  o  === TEXTO ===
        if (stripped.startswith("[") and stripped.endswith("]")) or \
           (stripped.startswith("===") and stripped.endswith("===")):
            titulo = stripped.strip("[]= ").strip()
            return (f'<span style="color:#60a5fa; font-weight:700; '
                    f'font-size:11px; letter-spacing:0.6px">▸ {self._escape(titulo)}</span>')

        # Controles / perfiles aprobados
        if stripped.startswith(("✓", "●")) or "[ACTIVO]" in stripped or stripped.startswith("[OK]"):
            return f'{indent_html}<span style="color:#86efac; font-weight:600">{esc}</span>'

        # Perfiles inactivos → línea compacta y tenue (sin su detalle, ver _log_pipeline)
        if stripped.startswith("○") or "[inactivo]" in stripped:
            return f'{indent_html}<span style="color:#4b5563; font-size:11px">{esc}</span>'

        # Controles / perfiles fallidos, bloqueos o advertencias
        if stripped.startswith(("✗", "⚠")) or stripped.startswith("[FALLA]") or stripped.startswith("ERROR"):
            return f'{indent_html}<span style="color:#fca5a5; font-weight:600">{esc}</span>'

        # Bullets de detalle anidado
        if stripped.startswith("•") or n_indent >= 4:
            return f'{indent_html}<span style="color:#9aa3b5">{esc}</span>'

        # Líneas clave: valor (ej. "ID: 123" o "Empresa : CGE")
        if ":" in stripped and len(stripped.split(":", 1)[0].strip()) <= 28:
            etiqueta, _, resto = stripped.partition(":")
            return (f'{indent_html}<span style="color:#7dd3fc">{self._escape(etiqueta)}:</span>'
                    f'<span style="color:#d4d8e2">{self._escape(resto)}</span>')

        # Línea genérica
        return f'{indent_html}<span style="color:#b8f5c0">{esc}</span>'

    def _scroll_console_to_bottom(self):
        self.console.moveCursor(QTextCursor.End)
        self.console.ensureCursorVisible()

    @staticmethod
    def _escape(texto: str) -> str:
        return (
            texto.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;")
        )

    def _limpiar_consola(self):
        self.console.clear()
        self._log_sistema("Consola limpiada.")

    def _on_ejecutar(self):
        n_referencia = self.input_referencia.text().strip()
        id_incidente = self.input_id_incidente.text().strip()

        if not n_referencia:
            QMessageBox.warning(self, "Falta información",
                                 "Debe ingresar el N° de referencia SEC.")
            return

        # Reset visual
        for card in self.stage_cards:
            card.set_estado(StageCard.PENDIENTE)
        self.progress_bar.setValue(0)
        self.estado_badge.set_estado(None)
        self.lbl_resumen.setText("Ejecutando pipeline…")
        self.btn_ejecutar.setEnabled(False)
        self.btn_ejecutar.setText("⏳ EJECUTANDO…")
        self.statusBar().showMessage("Ejecutando pipeline…")

        self._log_sistema(f"Iniciando pipeline — N° referencia: {n_referencia} "
                           f"| ID incidente (Excel): {id_incidente or '(usar el de SEC)'}")

        self.worker = PipelineWorker(
            n_referencia=n_referencia,
            id_incidente=id_incidente,
            id_pendientes=ID_PENDIENTES,
            excel_path=EXCEL_PATH,
        )
        self.worker.log_emitted.connect(self._log_pipeline)
        self.worker.stage_started.connect(self._on_stage_started)
        self.worker.stage_finished.connect(self._on_stage_finished)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.finished_error.connect(self._on_finished_error)
        self.worker.start()

    def _on_stage_started(self, indice: int, nombre: str):
        self.stage_cards[indice].set_estado(StageCard.EN_CURSO)
        self.statusBar().showMessage(f"Ejecutando etapa {indice + 1}/3 — {nombre}…")
        self._log_sistema(f"Etapa {indice + 1}/3 iniciada: {nombre}", "INFO")

    def _on_stage_finished(self, indice: int):
        self.stage_cards[indice].set_estado(StageCard.COMPLETADA)
        self.progress_bar.setValue(indice + 1)
        self._log_sistema(f"Etapa {indice + 1}/3 completada.", "OK")

    def _on_finished_ok(self, resultado: dict):
        self.btn_ejecutar.setEnabled(True)
        self.btn_ejecutar.setText("▶  EJECUTAR PIPELINE")
        self.statusBar().showMessage("Pipeline completado correctamente.")
        self._log_sistema("Pipeline completado correctamente.", "OK")

        caso          = resultado["caso"]
        clasificacion = resultado["clasificacion"]
        reporte       = resultado["lector"]

        estado = reporte.get("estado")
        self.estado_badge.set_estado(estado)

        tipo_cod    = reporte.get("tipologia", "N/A")
        tipo_nombre = TIPOLOGIA_NOMBRE.get(tipo_cod, "")
        confianza   = clasificacion.get("factor_de_confianza", 0)
        perfil      = reporte.get("perfil_principal") or "—"
        modif       = reporte.get("modificadores", [])
        alertas     = reporte.get("alertas", [])
        sev         = reporte.get("severidad", 0)
        banda       = reporte.get("banda_severidad", "Sin señal")
        causal      = reporte.get("causal_derivacion", "")
        texto_pre   = reporte.get("texto_prerresolucion", "")

        html = []
        html.append(f"<b>ID de incidente:</b> {caso.get('id', 'N/A')}")
        html.append(f"<b>N° referencia:</b> {caso.get('n_referencia', 'N/A')}")
        html.append(f"<b>Fecha reclamo:</b> {caso.get('fecha', 'N/A')}")
        html.append(f"<b>Empresa:</b> {caso.get('empresa', 'N/A')}")
        html.append("<br>")
        html.append(f"<b>Clasificación IA:</b> {clasificacion.get('clasificacion', 'N/A')}")
        html.append(f"<b>Confianza:</b> {round(confianza * 100)}%")
        html.append(f"<b>Tipología:</b> {tipo_cod} — {tipo_nombre}")
        html.append("<br>")
        html.append(f"<b>Perfil principal:</b> {perfil}")
        html.append(f"<b>Modificadores:</b> {', '.join(modif) if modif else '—'}")
        html.append(f"<b>Alertas:</b> {', '.join(alertas) if alertas else '—'}")
        html.append(f"<b>Severidad:</b> {sev} pts ({banda})")

        if causal:
            html.append("<br><b>Causal de derivación:</b>")
            for parte in causal.split("; "):
                parte = parte.strip()
                if parte:
                    html.append(f"&nbsp;&nbsp;• {parte}")

        if texto_pre:
            html.append("<br><b>Prerresolución automática:</b>")
            html.append(f"<i>{texto_pre}</i>")

        self.lbl_resumen.setText("<br>".join(html))

    def _on_finished_error(self, etapa: str, mensaje: str):
        self.btn_ejecutar.setEnabled(True)
        self.btn_ejecutar.setText("▶  EJECUTAR PIPELINE")
        self.statusBar().showMessage(f"Error en etapa: {etapa}")

        # Marcar como error la tarjeta de la etapa en curso
        for card in self.stage_cards:
            if card.estado == StageCard.EN_CURSO:
                card.set_estado(StageCard.ERROR)

        self._log_sistema(f"ERROR en {etapa}: {mensaje}", "ERROR")
        self.estado_badge.setText("ERROR")
        self.estado_badge.setStyleSheet(
            f"background-color: {COLOR_RED}22; color: {COLOR_RED}; "
            f"border: 1.5px solid {COLOR_RED}; border-radius: 8px; "
            f"font-weight: 700; font-size: 15px;"
        )
        self.lbl_resumen.setText(
            f"<b style='color:{COLOR_RED}'>Error durante la ejecución</b><br><br>"
            f"<b>Etapa:</b> {etapa}<br>"
            f"<b>Detalle:</b><br>{self._escape(mensaje).replace(chr(10), '<br>')}"
        )

        QMessageBox.critical(self, f"Error en {etapa}", mensaje[:600])

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()