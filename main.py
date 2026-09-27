# -*- coding: utf-8 -*-
"""
RPA - Atualização de Arquivos Excel (Refresh All)

Abre cada arquivo configurado em ARCHIVES, executa RefreshAll,
aguarda a conclusão das atualizações, compara a quantidade de linhas
por planilha, salva e fecha o arquivo. A execução é monitorada por
uma interface gráfica e gera log detalhado em ./logs.
"""

from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
import time
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

APP_TITLE = "Atualização de Arquivos Excel"
APP_SUBTITLE = (
    "RefreshAll monitorado, validação de dados e acompanhamento das fontes corporativas"
)

# ─── DEPENDÊNCIAS WINDOWS / EXCEL ────────────────────────────────────────
try:
    import pythoncom
    import win32com.client as win32
    PYWIN32_ERROR = None
except ImportError as exc:
    pythoncom = None
    win32 = None
    PYWIN32_ERROR = exc


# ─── CONFIGURAÇÕES GERAIS ────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOG_DIR / f"rpa_excel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)
LOGGER = logging.getLogger(__name__)

# Caminho raiz do compartilhamento de rede onde as planilhas estão
# hospedadas. Ajuste o valor abaixo (ou defina a variável de ambiente
# RPA_SHARE_ROOT) para o caminho real do seu ambiente. Manter esse tipo de
# informação fora do código-fonte evita expor a estrutura interna da rede
# em repositórios públicos.
SHARE_ROOT = os.environ.get(
    "RPA_SHARE_ROOT",
    r"\\SEU-SERVIDOR\planilha\Inteligencia_Comercial",
)

ARCHIVES = [
    rf"{SHARE_ROOT}\Automacoes\Painel Diário\Painel Diario.xlsx",
    rf"{SHARE_ROOT}\Automacoes\Passagens - CD\Passagens - CD.xlsx",
    # Extensão corrigida para .xlsb (anteriormente configurada como .xlsx).
    rf"{SHARE_ROOT}\Automacoes\Quinta do Óleo\Promoção Quinta do Óleo - Teste.xlsb",
    rf"{SHARE_ROOT}\BI - Pós Vendas\2026_2R_Farol_Indicadores_REGER (Calculadora).xlsx",
    rf"{SHARE_ROOT}\BI - Pós Vendas\2026_4R_Farol_Indicadores_REGER (Calculadora)_V2.xlsx",
    rf"{SHARE_ROOT}\BI - Pós Vendas\Pós Vendas - Fato.xlsx",
]

# Tempo máximo real de espera pelo RefreshAll de cada arquivo.
TIMEOUT = 600

# Intervalo de verificação do estado de atualização.
POLL_INTERVAL = 0.5

# Número de verificações consecutivas sem atualização/calculação ativa
# antes de considerar o RefreshAll concluído.
STABLE_CHECKS_REQUIRED = 3


# ─── EXCEÇÕES / UTILITÁRIOS ──────────────────────────────────────────────
class ExecutionCancelled(Exception):
    """Sinaliza interrupção solicitada pelo usuário."""


def now_time() -> str:
    return datetime.now().strftime("%H:%M:%S")


def format_duration(seconds: float) -> str:
    seconds = max(0, int(seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def validate_archive(path: str) -> tuple[bool, str]:
    p = Path(path)

    if not p.parent.exists():
        return False, f"Diretório não encontrado: {p.parent}"

    if not p.exists():
        return False, f"Arquivo não encontrado: {p}"

    if not p.is_file():
        return False, f"O caminho não aponta para um arquivo: {p}"

    return True, str(p)


def count_rows_by_sheet(workbook) -> dict[str, int]:
    """
    Retorna a última linha efetivamente usada de cada planilha.

    Usa Cells.Find em vez de UsedRange.Rows.Count para evitar que
    formatação residual faça a planilha parecer ter milhares de linhas.
    """
    counts: dict[str, int] = {}

    for ws in workbook.Worksheets:
        try:
            # Constantes Excel:
            # xlFormulas = -4123 | xlByRows = 1 | xlPrevious = 2
            last_cell = ws.Cells.Find(
                What="*",
                After=ws.Cells(1, 1),
                LookIn=-4123,
                SearchOrder=1,
                SearchDirection=2,
            )
            counts[ws.Name] = int(last_cell.Row) if last_cell is not None else 0
        except Exception:
            counts[ws.Name] = -1

    return counts


def format_rows(data: dict[str, int]) -> str:
    if not data:
        return "Nenhuma planilha identificada"
    return ", ".join(f"{name}: {qty}" for name, qty in data.items())


def row_deltas(before: dict[str, int], after: dict[str, int]) -> list[str]:
    """Inclui abas existentes antes, depois ou criadas/removidas no refresh."""
    lines = []
    sheet_names = list(dict.fromkeys([*before.keys(), *after.keys()]))

    for name in sheet_names:
        a = before.get(name)
        d = after.get(name)

        if a is None:
            lines.append(f"  - {name}: criada durante a atualização -> {d}")
            continue

        if d is None:
            lines.append(f"  - {name}: removida durante a atualização (antes: {a})")
            continue

        variation = "N/D" if (a == -1 or d == -1) else str(d - a)
        lines.append(f"  - {name}: {a} -> {d} (delta: {variation})")

    return lines


def workbook_has_active_refresh(workbook) -> bool:
    """
    Tenta detectar consultas/conexões que ainda estejam atualizando.
    A API do Excel não é uniforme entre ODBC, OLEDB, QueryTables e ListObjects,
    por isso cada fonte é verificada defensivamente.
    """
    # Conexões do workbook.
    try:
        for connection in workbook.Connections:
            try:
                if bool(connection.OLEDBConnection.Refreshing):
                    return True
            except Exception:
                pass

            try:
                if bool(connection.ODBCConnection.Refreshing):
                    return True
            except Exception:
                pass
    except Exception:
        pass

    # QueryTables e ListObjects por planilha.
    try:
        for ws in workbook.Worksheets:
            try:
                for query_table in ws.QueryTables:
                    try:
                        if bool(query_table.Refreshing):
                            return True
                    except Exception:
                        pass
            except Exception:
                pass

            try:
                for list_object in ws.ListObjects:
                    try:
                        query_table = list_object.QueryTable
                        if query_table is not None and bool(query_table.Refreshing):
                            return True
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

    return False


def wait_for_refresh(workbook, excel_app, cancel_event: threading.Event) -> float:
    """
    Aguarda o RefreshAll terminar e aplica TIMEOUT de verdade.

    O script anterior apenas verificava o tempo depois da conclusão,
    portanto TIMEOUT era um aviso, não um limite operacional.
    """
    start = time.perf_counter()
    stable_checks = 0

    while True:
        if cancel_event.is_set():
            raise ExecutionCancelled("Execução interrompida pelo usuário.")

        elapsed = time.perf_counter() - start
        if elapsed > TIMEOUT:
            raise TimeoutError(
                f"A atualização excedeu o limite de {TIMEOUT} segundos."
            )

        refreshing = workbook_has_active_refresh(workbook)

        try:
            # xlDone = 0
            calculation_done = int(excel_app.CalculationState) == 0
        except Exception:
            calculation_done = True

        if not refreshing and calculation_done:
            stable_checks += 1
            if stable_checks >= STABLE_CHECKS_REQUIRED:
                return elapsed
        else:
            stable_checks = 0

        time.sleep(POLL_INTERVAL)


# ─── INTERFACE ───────────────────────────────────────────────────────────
class ExcelRefreshApp(tk.Tk):
    BG = "#101114"
    PANEL = "#17191F"
    PANEL_DARK = "#13161C"
    CARD = "#1B1F27"
    CARD_ALT = "#232833"
    LOG_BG = "#0E1015"

    ACCENT = "#E10600"
    TEXT = "#FFFFFF"
    TEXT_SOFT = "#D5D7DC"
    TEXT_MUTED = "#8F96A3"
    BORDER = "#2A2D34"

    SUCCESS = "#22C55E"
    WARNING = "#F59E0B"
    ERROR = "#FF4D4F"
    INACTIVE = "#4B5563"

    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1280x790")
        self.minsize(1120, 700)
        self.configure(bg=self.BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread = None
        self.running = False
        self.started_at = None

        self.results = []

        self.status_var = tk.StringVar(value="Sistema pronto para iniciar")
        self.detail_var = tk.StringVar(
            value="Execute a rotina para atualizar os arquivos Excel configurados."
        )
        self.badge_var = tk.StringVar(value="RPA AGUARDANDO")
        self.progress_text_var = tk.StringVar(value=f"0 / {len(ARCHIVES)}")
        self.percent_var = tk.StringVar(value="0%")
        self.timer_var = tk.StringVar(value="00:00")
        self.current_file_var = tk.StringVar(value="Nenhum arquivo em execução")
        self.success_var = tk.StringVar(value="0")
        self.failure_var = tk.StringVar(value="0")

        self.step_dots = []
        self.step_labels = []

        self._configure_styles()
        self._build_ui()

        self.after(100, self._process_events)
        self.after(500, self._update_timer)

        if PYWIN32_ERROR is not None:
            self.after(250, self._show_dependency_warning)

    # ─── Estilos ─────────────────────────────────────────────────────────
    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=self.BORDER,
            background=self.ACCENT,
            darkcolor=self.ACCENT,
            lightcolor=self.ACCENT,
            bordercolor=self.BORDER,
            thickness=18,
        )

        style.configure(
            "Primary.TButton",
            background=self.ACCENT,
            foreground=self.TEXT,
            borderwidth=0,
            focusthickness=0,
            padding=(14, 9),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", "#5B2020"),
                ("active", "#B80500"),
                ("pressed", "#930400"),
            ],
            foreground=[("disabled", "#B9B9B9")],
        )

        style.configure(
            "Secondary.TButton",
            background=self.BORDER,
            foreground=self.TEXT,
            borderwidth=0,
            focusthickness=0,
            padding=(12, 9),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#3A3E47"), ("pressed", "#20232A")],
        )

        style.configure(
            "Danger.TButton",
            background="#3A2022",
            foreground="#FFD7D7",
            borderwidth=0,
            focusthickness=0,
            padding=(12, 9),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Danger.TButton",
            background=[
                ("disabled", "#282022"),
                ("active", "#5A2629"),
                ("pressed", "#2B1718"),
            ],
            foreground=[("disabled", "#777777")],
        )

    # ─── Construção visual ───────────────────────────────────────────────
    def _build_ui(self):
        header = tk.Frame(self, bg=self.BG)
        header.pack(fill="x", padx=20, pady=(18, 10))

        self.badge_label = tk.Label(
            header,
            textvariable=self.badge_var,
            bg=self.WARNING,
            fg=self.TEXT,
            font=("Segoe UI", 9, "bold"),
            padx=11,
            pady=5,
        )
        self.badge_label.pack(anchor="w", pady=(0, 9))

        tk.Label(
            header,
            text=APP_TITLE,
            bg=self.BG,
            fg=self.TEXT,
            font=("Segoe UI", 25, "bold"),
        ).pack(anchor="w")

        tk.Label(
            header,
            text=APP_SUBTITLE,
            bg=self.BG,
            fg=self.TEXT_SOFT,
            font=("Segoe UI", 12),
        ).pack(anchor="w", pady=(4, 0))

        controls = tk.Frame(self, bg=self.BG)
        controls.pack(fill="x", padx=20, pady=(0, 10))

        self.start_button = ttk.Button(
            controls,
            text="Executar atualização",
            command=self.start_execution,
            style="Primary.TButton",
        )
        self.start_button.pack(side="left")

        self.cancel_button = ttk.Button(
            controls,
            text="Interromper",
            command=self.cancel_execution,
            style="Danger.TButton",
            state="disabled",
        )
        self.cancel_button.pack(side="left", padx=(7, 0))

        ttk.Button(
            controls,
            text="Abrir pasta de logs",
            command=self.open_log_folder,
            style="Secondary.TButton",
        ).pack(side="left", padx=(7, 0))

        body = tk.Frame(self, bg=self.BG)
        body.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        left = tk.Frame(body, bg=self.PANEL, width=405)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        right = tk.Frame(body, bg=self.PANEL_DARK)
        right.pack(side="right", fill="both", expand=True, padx=(14, 0))

        # Pipeline de arquivos
        tk.Label(
            left,
            text="Pipeline operacional",
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w", padx=16, pady=(16, 4))

        tk.Label(
            left,
            text=(
                f"{len(ARCHIVES)} arquivos configurados. Cada arquivo é aberto, "
                "atualizado, validado, salvo e fechado individualmente."
            ),
            bg=self.PANEL,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
            wraplength=365,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(0, 12))

        steps_wrap = tk.Frame(left, bg=self.PANEL)
        steps_wrap.pack(fill="both", expand=True, padx=16)

        for index, archive in enumerate(ARCHIVES):
            row = tk.Frame(steps_wrap, bg=self.PANEL)
            row.pack(fill="x", pady=5)

            dot = tk.Label(
                row,
                text="●",
                bg=self.PANEL,
                fg=self.INACTIVE,
                font=("Segoe UI", 10, "bold"),
                width=2,
            )
            dot.pack(side="left", anchor="n")

            label = tk.Label(
                row,
                text=f"{index + 1:02d}. {Path(archive).name}",
                bg=self.PANEL,
                fg=self.TEXT_MUTED,
                font=("Segoe UI", 9),
                wraplength=330,
                justify="left",
                anchor="w",
            )
            label.pack(side="left", fill="x", expand=True)

            self.step_dots.append(dot)
            self.step_labels.append(label)

        info_box = tk.Frame(left, bg=self.CARD_ALT)
        info_box.pack(fill="x", padx=16, pady=16)

        tk.Label(
            info_box,
            text="Escopo da rotina",
            bg=self.CARD_ALT,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=12, pady=(10, 2))

        tk.Label(
            info_box,
            text=(
                "RefreshAll via Microsoft Excel, comparação da última linha usada "
                "por aba e salvamento apenas após atualização concluída."
            ),
            bg=self.CARD_ALT,
            fg=self.TEXT_SOFT,
            font=("Segoe UI", 9),
            wraplength=340,
            justify="left",
        ).pack(anchor="w", padx=12)

        tk.Label(
            info_box,
            text=f"Log: logs\\{LOG_FILE.name}",
            bg=self.CARD_ALT,
            fg=self.TEXT_MUTED,
            font=("Consolas", 8),
            wraplength=340,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(7, 10))

        # Status
        status_card = tk.Frame(right, bg=self.CARD)
        status_card.pack(fill="x", padx=16, pady=(16, 10))

        tk.Label(
            status_card,
            text="Status da execução",
            bg=self.CARD,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=14, pady=(12, 2))

        tk.Label(
            status_card,
            textvariable=self.status_var,
            bg=self.CARD,
            fg=self.TEXT,
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w", padx=14)

        tk.Label(
            status_card,
            textvariable=self.detail_var,
            bg=self.CARD,
            fg=self.ACCENT,
            font=("Segoe UI", 10, "bold"),
            wraplength=760,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(4, 12))

        # KPIs
        metrics = tk.Frame(right, bg=self.PANEL_DARK)
        metrics.pack(fill="x", padx=16, pady=(0, 10))

        metric_items = [
            ("PROGRESSO", self.progress_text_var, self.TEXT),
            ("SUCESSOS", self.success_var, self.SUCCESS),
            ("FALHAS", self.failure_var, self.ERROR),
            ("TEMPO", self.timer_var, self.WARNING),
        ]

        for index, (title, variable, color) in enumerate(metric_items):
            card = tk.Frame(metrics, bg=self.CARD)
            card.grid(
                row=0,
                column=index,
                sticky="nsew",
                padx=(0 if index == 0 else 5, 0 if index == len(metric_items) - 1 else 5),
            )
            metrics.grid_columnconfigure(index, weight=1)

            tk.Label(
                card,
                text=title,
                bg=self.CARD,
                fg=self.TEXT_MUTED,
                font=("Segoe UI", 8, "bold"),
            ).pack(anchor="w", padx=12, pady=(10, 2))

            tk.Label(
                card,
                textvariable=variable,
                bg=self.CARD,
                fg=color,
                font=("Segoe UI", 18, "bold"),
            ).pack(anchor="w", padx=12, pady=(0, 10))

        # Progresso
        progress_card = tk.Frame(right, bg=self.CARD)
        progress_card.pack(fill="x", padx=16, pady=(0, 10))

        progress_head = tk.Frame(progress_card, bg=self.CARD)
        progress_head.pack(fill="x", padx=14, pady=(12, 6))

        tk.Label(
            progress_head,
            text="Arquivo atual",
            bg=self.CARD,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(side="left")

        tk.Label(
            progress_head,
            textvariable=self.current_file_var,
            bg=self.CARD,
            fg=self.TEXT_SOFT,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right")

        self.progress = ttk.Progressbar(
            progress_card,
            style="Accent.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate",
            maximum=max(1, len(ARCHIVES)),
            value=0,
        )
        self.progress.pack(fill="x", padx=14, pady=(0, 12))

        # Log visual
        log_card = tk.Frame(right, bg=self.CARD)
        log_card.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        log_head = tk.Frame(log_card, bg=self.CARD)
        log_head.pack(fill="x", padx=14, pady=(12, 8))

        tk.Label(
            log_head,
            text="Log operacional",
            bg=self.CARD,
            fg=self.TEXT,
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")

        tk.Label(
            log_head,
            text="Saída em tempo real",
            bg=self.CARD,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(side="right")

        log_frame = tk.Frame(log_card, bg=self.LOG_BG)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.log_text = tk.Text(
            log_frame,
            bg=self.LOG_BG,
            fg="#D8DCE5",
            insertbackground=self.TEXT,
            font=("Consolas", 9),
            relief="flat",
            borderwidth=0,
            padx=10,
            pady=8,
            wrap="word",
            state="disabled",
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.tag_configure("info", foreground="#D8DCE5")
        self.log_text.tag_configure("success", foreground="#8DE4A3")
        self.log_text.tag_configure("warning", foreground="#F5C96B")
        self.log_text.tag_configure("error", foreground="#FF9294")
        self.log_text.tag_configure("accent", foreground="#FF7A75")

        footer = tk.Frame(self, bg=self.BG)
        footer.pack(fill="x", padx=20, pady=(0, 14))

        tk.Label(
            footer,
            text=(
                f"Timeout por arquivo: {TIMEOUT}s | "
                "Falhas são registradas e a rotina segue para o próximo arquivo"
            ),
            bg=self.BG,
            fg=self.TEXT_MUTED,
            font=("Segoe UI", 9),
        ).pack(side="left")

        ttk.Button(
            footer,
            text="Fechar",
            command=self._on_close,
            style="Secondary.TButton",
        ).pack(side="right")

    # ─── Feedback visual ─────────────────────────────────────────────────
    def _show_dependency_warning(self):
        self.status_var.set("Dependência ausente")
        self.detail_var.set("O pacote pywin32 é necessário para controlar o Microsoft Excel.")
        self._set_badge("DEPENDÊNCIA AUSENTE", self.ERROR)
        self.start_button.configure(state="disabled")
        self._append_log(
            "pywin32 não está instalado. Execute: pip install pywin32",
            "error",
        )
        messagebox.showerror(
            "Dependência ausente",
            "pywin32 não está instalado.\n\nExecute:\npip install pywin32",
        )

    def _set_badge(self, text: str, color: str) -> None:
        self.badge_var.set(text)
        self.badge_label.configure(bg=color)

    def _set_step(self, index: int, state: str) -> None:
        palettes = {
            "inactive": (self.INACTIVE, self.TEXT_MUTED),
            "running": (self.ACCENT, self.TEXT),
            "done": (self.SUCCESS, "#C9F4D4"),
            "error": (self.ERROR, "#FFD1D2"),
            "warning": (self.WARNING, "#FFE3A3"),
        }
        dot_color, text_color = palettes[state]
        self.step_dots[index].configure(fg=dot_color)
        self.step_labels[index].configure(fg=text_color)

    def _reset_ui(self):
        self.results = []
        for index in range(len(ARCHIVES)):
            self._set_step(index, "inactive")

        self.progress.configure(value=0)
        self.progress_text_var.set(f"0 / {len(ARCHIVES)}")
        self.percent_var.set("0%")
        self.timer_var.set("00:00")
        self.current_file_var.set("Nenhum arquivo em execução")
        self.success_var.set("0")
        self.failure_var.set("0")

    def _append_log(self, message, tag="info"):
        timestamp = now_time()
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ─── Execução ────────────────────────────────────────────────────────
    def start_execution(self):
        if self.running or PYWIN32_ERROR is not None:
            return

        self.running = True
        self.cancel_event.clear()
        self.started_at = datetime.now()
        self._reset_ui()

        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")

        self.status_var.set("Preparando execução")
        self.detail_var.set("Validando os arquivos configurados antes de iniciar o Excel.")
        self._set_badge("RPA INICIANDO", self.WARNING)

        self._append_log("=" * 78, "accent")
        self._append_log("INÍCIO DO RPA DE ATUALIZAÇÃO DE ARQUIVOS EXCEL", "accent")
        self._append_log(f"Log: {LOG_FILE}", "info")

        LOGGER.info("=" * 80)
        LOGGER.info("INÍCIO DO RPA DE ATUALIZAÇÃO DE ARQUIVOS EXCEL")
        LOGGER.info("Arquivos configurados: %s", len(ARCHIVES))
        LOGGER.info("=" * 80)

        self.worker_thread = threading.Thread(
            target=self._run_pipeline,
            daemon=True,
        )
        self.worker_thread.start()

    def _run_pipeline(self):
        if pythoncom is None or win32 is None:
            self.events.put(("fatal_error", "pywin32 não está disponível."))
            return

        pythoncom.CoInitialize()
        excel_app = None

        try:
            # Preflight completo.
            invalid_files = []
            for index, archive in enumerate(ARCHIVES):
                if self.cancel_event.is_set():
                    raise ExecutionCancelled

                valid, detail = validate_archive(archive)
                if not valid:
                    invalid_files.append((index, detail))

            if invalid_files:
                self.events.put(("preflight_error", invalid_files))
                return

            self.events.put(("preflight_ok", None))

            LOGGER.info("Inicializando Microsoft Excel em segundo plano.")
            excel_app = win32.DispatchEx("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.AskToUpdateLinks = False
            excel_app.ScreenUpdating = False

            for index, archive in enumerate(ARCHIVES):
                if self.cancel_event.is_set():
                    raise ExecutionCancelled

                result = self._process_archive(index, archive, excel_app)
                self.results.append(result)
                self.events.put(("file_result", index, result))

            self.events.put(("pipeline_done", self.results))

        except ExecutionCancelled:
            LOGGER.warning("Execução interrompida pelo usuário.")
            self.events.put(("cancelled", None))

        except Exception as exc:
            LOGGER.exception("Erro geral na execução do RPA.")
            self.events.put(("fatal_error", str(exc), traceback.format_exc()))

        finally:
            if excel_app is not None:
                try:
                    excel_app.Quit()
                    LOGGER.info("Aplicação Excel finalizada.")
                except Exception as exc:
                    LOGGER.warning("Erro ao finalizar o Excel: %s", exc)

            pythoncom.CoUninitialize()

    def _process_archive(self, index, path, excel_app):
        name = Path(path).name
        workbook = None
        start = time.perf_counter()

        result = {
            "arquivo": name,
            "status": "FALHA",
            "erro": None,
            "tempo_atualizacao_s": None,
            "linhas_iniciais": {},
            "linhas_finais": {},
        }

        self.events.put(("file_start", index, path))

        LOGGER.info("-" * 80)
        LOGGER.info("ARQUIVO: %s", name)
        LOGGER.info("CAMINHO: %s", path)

        try:
            if self.cancel_event.is_set():
                raise ExecutionCancelled

            LOGGER.info("Abrindo arquivo: %s", path)
            workbook = excel_app.Workbooks.Open(
                path,
                UpdateLinks=0,
                ReadOnly=False,
            )

            if bool(workbook.ReadOnly):
                raise PermissionError(
                    "O arquivo foi aberto como somente leitura. "
                    "Ele pode estar em uso por outro usuário/processo."
                )

            before = count_rows_by_sheet(workbook)
            result["linhas_iniciais"] = before
            LOGGER.info("Linhas antes: %s", format_rows(before))
            self.events.put(("file_log", f"Linhas antes: {format_rows(before)}", "info"))

            self.events.put(("file_log", "Iniciando RefreshAll...", "accent"))
            LOGGER.info("Iniciando RefreshAll.")

            workbook.RefreshAll()
            refresh_elapsed = wait_for_refresh(
                workbook,
                excel_app,
                self.cancel_event,
            )

            result["tempo_atualizacao_s"] = round(refresh_elapsed, 2)
            LOGGER.info("RefreshAll concluído em %.2f s.", refresh_elapsed)

            after = count_rows_by_sheet(workbook)
            result["linhas_finais"] = after
            LOGGER.info("Linhas depois: %s", format_rows(after))

            for line in row_deltas(before, after):
                LOGGER.info(line)
                self.events.put(("file_log", line, "info"))

            if self.cancel_event.is_set():
                raise ExecutionCancelled

            self.events.put(("file_log", "Salvando arquivo...", "accent"))
            workbook.Save()

            result["status"] = "SUCESSO"
            self.events.put(
                (
                    "file_log",
                    f"Arquivo salvo com sucesso. Refresh: {format_duration(refresh_elapsed)}",
                    "success",
                )
            )
            LOGGER.info("Arquivo salvo com sucesso.")

        except ExecutionCancelled:
            result["status"] = "CANCELADO"
            result["erro"] = "Execução interrompida pelo usuário."
            raise

        except TimeoutError as exc:
            result["erro"] = str(exc)
            LOGGER.error("Timeout: %s", exc)
            self.events.put(("file_log", f"TIMEOUT: {exc}", "error"))

        except Exception as exc:
            result["erro"] = str(exc)
            LOGGER.exception("Erro ao processar %s", name)
            self.events.put(("file_log", f"ERRO: {exc}", "error"))

        finally:
            if workbook is not None:
                try:
                    # SaveChanges=False evita salvar parcialmente quando houve falha.
                    workbook.Close(SaveChanges=False)
                    LOGGER.info("Arquivo fechado: %s", name)
                except Exception as exc:
                    LOGGER.warning("Erro ao fechar o arquivo %s: %s", name, exc)

        result["tempo_total_s"] = round(time.perf_counter() - start, 2)
        return result

    def cancel_execution(self):
        if not self.running:
            return

        if not messagebox.askyesno(
            "Interromper execução",
            "Deseja realmente interromper a rotina em andamento?",
        ):
            return

        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self._set_badge("CANCELANDO", self.WARNING)
        self.status_var.set("Interrompendo execução")
        self.detail_var.set(
            "A solicitação foi registrada e será aplicada assim que o Excel liberar o estágio atual."
        )
        self._append_log("Solicitação de cancelamento recebida.", "warning")

    # ─── Eventos do worker -> Tkinter ────────────────────────────────────
    def _process_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._process_events)

    def _handle_event(self, event):
        event_type = event[0]

        if event_type == "preflight_ok":
            self.status_var.set("Validação concluída")
            self.detail_var.set("Todos os arquivos foram localizados. Inicializando o Excel.")
            self._set_badge("RPA EM EXECUÇÃO", self.ACCENT)
            self._append_log("Validação prévia concluída com sucesso.", "success")

        elif event_type == "preflight_error":
            invalid_files = event[1]
            self.status_var.set("Falha na validação inicial")
            self.detail_var.set(
                f"{len(invalid_files)} arquivo(s) não puderam ser localizado(s)."
            )
            self._set_badge("RPA COM ERRO", self.ERROR)

            for index, detail in invalid_files:
                self._set_step(index, "error")
                self._append_log(detail, "error")
                LOGGER.error(detail)

            self._finish_running_state()

        elif event_type == "file_start":
            _, index, path = event
            self._set_step(index, "running")
            self.status_var.set(f"Atualizando arquivo {index + 1} de {len(ARCHIVES)}")
            self.detail_var.set(Path(path).name)
            self.current_file_var.set(Path(path).name)
            self._append_log(f"Abrindo e atualizando: {Path(path).name}", "accent")

        elif event_type == "file_log":
            _, message, tag = event
            self._append_log(message, tag)

        elif event_type == "file_result":
            _, index, result = event

            if result["status"] == "SUCESSO":
                self._set_step(index, "done")
            else:
                self._set_step(index, "error")

            processed = index + 1
            successes = sum(1 for item in self.results if item["status"] == "SUCESSO")
            failures = sum(1 for item in self.results if item["status"] == "FALHA")

            percent = round((processed / len(ARCHIVES)) * 100) if ARCHIVES else 100

            self.progress.configure(value=processed)
            self.progress_text_var.set(f"{processed} / {len(ARCHIVES)}")
            self.percent_var.set(f"{percent}%")
            self.success_var.set(str(successes))
            self.failure_var.set(str(failures))

            if result["status"] == "SUCESSO":
                self._append_log(
                    f"Concluído: {result['arquivo']} | "
                    f"refresh={format_duration(result['tempo_atualizacao_s'] or 0)}",
                    "success",
                )
            else:
                self._append_log(
                    f"Falha: {result['arquivo']} | {result['erro']}",
                    "error",
                )

        elif event_type == "pipeline_done":
            results = event[1]
            successes = sum(1 for item in results if item["status"] == "SUCESSO")
            failures = sum(1 for item in results if item["status"] == "FALHA")
            elapsed = (
                (datetime.now() - self.started_at).total_seconds()
                if self.started_at
                else 0
            )

            self.progress.configure(value=len(ARCHIVES))
            self.progress_text_var.set(f"{len(ARCHIVES)} / {len(ARCHIVES)}")
            self.percent_var.set("100%")
            self.success_var.set(str(successes))
            self.failure_var.set(str(failures))
            self.current_file_var.set("Rotina finalizada")

            if failures == 0:
                self.status_var.set("RPA finalizado com sucesso")
                self.detail_var.set(
                    f"{successes} arquivo(s) atualizados em {format_duration(elapsed)}."
                )
                self._set_badge("RPA CONCLUÍDO", self.SUCCESS)
            else:
                self.status_var.set("RPA finalizado com ocorrências")
                self.detail_var.set(
                    f"Sucessos: {successes} | Falhas: {failures} | "
                    f"Tempo: {format_duration(elapsed)}"
                )
                self._set_badge("RPA COM OCORRÊNCIAS", self.WARNING)

            self._write_summary_to_log(results, elapsed)
            self._finish_running_state()

        elif event_type == "cancelled":
            self.status_var.set("Execução interrompida")
            self.detail_var.set("A rotina foi cancelada pelo usuário.")
            self._set_badge("RPA INTERROMPIDO", self.WARNING)
            self.current_file_var.set("Execução cancelada")
            self._append_log("RPA interrompido pelo usuário.", "warning")
            self._finish_running_state()

        elif event_type == "fatal_error":
            detail = event[1]
            traceback_text = event[2] if len(event) > 2 else ""

            self.status_var.set("Falha geral na execução")
            self.detail_var.set(detail)
            self._set_badge("RPA COM ERRO", self.ERROR)
            self._append_log(detail, "error")

            if traceback_text:
                LOGGER.error(traceback_text)

            self._finish_running_state()

    def _write_summary_to_log(self, results, elapsed):
        successes = sum(1 for item in results if item["status"] == "SUCESSO")
        failures = sum(1 for item in results if item["status"] == "FALHA")

        LOGGER.info("=" * 80)
        LOGGER.info("RESUMO GERAL")
        LOGGER.info("Duração: %s", format_duration(elapsed))
        LOGGER.info("Sucessos: %s/%s", successes, len(results))
        LOGGER.info("Falhas: %s/%s", failures, len(results))

        for result in results:
            LOGGER.info("[%s] %s", result["status"], result["arquivo"])
            if result["status"] == "SUCESSO":
                LOGGER.info(
                    "  Refresh: %.2f s",
                    result["tempo_atualizacao_s"] or 0,
                )
                LOGGER.info(
                    "  Linhas iniciais: %s",
                    format_rows(result["linhas_iniciais"]),
                )
                LOGGER.info(
                    "  Linhas finais: %s",
                    format_rows(result["linhas_finais"]),
                )
            else:
                LOGGER.info("  Erro: %s", result["erro"])

        LOGGER.info("=" * 80)

    def _finish_running_state(self):
        self.running = False
        self.start_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")

    # ─── Utilidades ──────────────────────────────────────────────────────
    def _update_timer(self):
        if self.running and self.started_at:
            elapsed = (datetime.now() - self.started_at).total_seconds()
            self.timer_var.set(format_duration(elapsed))

        self.after(500, self._update_timer)

    def open_log_folder(self):
        try:
            if os.name == "nt":
                os.startfile(LOG_DIR)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(LOG_DIR)])
            else:
                subprocess.Popen(["xdg-open", str(LOG_DIR)])
        except Exception as exc:
            messagebox.showerror(
                "Abrir pasta de logs",
                f"Não foi possível abrir a pasta:\n{exc}",
            )

    def _on_close(self):
        if self.running:
            close = messagebox.askyesno(
                "RPA em execução",
                "Existe uma atualização em andamento. Deseja solicitar interrupção e fechar?",
            )
            if not close:
                return

            self.cancel_event.set()

        self.destroy()


# ─── EXECUÇÃO PRINCIPAL ──────────────────────────────────────────────────
def main() -> None:
    app = ExcelRefreshApp()
    app.mainloop()


if __name__ == "__main__":
    main()
