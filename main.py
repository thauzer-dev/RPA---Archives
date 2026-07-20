# -*- coding: utf-8 -*-
"""
RPA - Atualização de Arquivos Excel (Refresh All)

Abre cada arquivo da lista `ARCHIVES`, executa Atualizar Tudo (RefreshAll),
registra variação de linhas por planilha, salva e fecha. Gera log detalhado.

Uso:
    python main.py
"""

import os 
import sys 
import io
import time 
import datetime 
import traceback 
from pathlib import Path  # CORRIGIDO: era "import pathlib as Path"

# Força stdout/stderr em UTF-8 para evitar UnicodeEncodeError no terminal CP1252
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

try:
    import win32com.client as win32
except ImportError:
    print("ERRO: pywin32 não está instalado. Execute: pip install pywin32")
    sys.exit(1)


# ─── Configurações ────────────────────────────────────────────────────────────

ARCHIVES = [
    r"\\192.168.4.16\planilha\Inteligência Comercial\Enzo Thauzer\Python\Automações\Painel Diário\Painel Diario.xlsx",
    r"\\192.168.4.16\planilha\Inteligência Comercial\Enzo Thauzer\Python\Automações\Margem Bruta - 2R\MB%_2R.xlsb",
    r"\\192.168.4.16\planilha\Inteligência Comercial\Enzo Thauzer\Python\Automações\Passagens - CD\Passagens - CD.xlsx",
    r"\\192.168.4.16\planilha\Inteligência Comercial\Enzo Thauzer\Python\Automações\Quinta do Óleo\Promoção Quinta do Óleo - Teste.xlsb"  # CORRIGIDO: era .xslx
]

# Tempo máximo (segundos) de espera pela atualização antes de emitir aviso
TIMEOUT = 600

# Pasta de destino do log (None = mesma pasta do script)
PASTA_LOG: Path | None = None

# Pasta base do script — usada como fallback para o log
_SCRIPT_DIR = Path(__file__).resolve().parent


# ─── Utilitários ──────────────────────────────────────────────────────────────

def _agora() -> str:
    """Retorna hora atual formatada para uso no log."""
    return datetime.datetime.now().strftime("%H:%M:%S")


def contar_linhas_planilhas(workbook) -> dict[str, int]:
    """Retorna {nome_planilha: linhas_usadas} para todas as abas do workbook."""
    contagem: dict[str, int] = {}
    for ws in workbook.Worksheets:
        try:
            contagem[ws.Name] = ws.UsedRange.Rows.Count
        except Exception:
            contagem[ws.Name] = -1  # indisponível
    return contagem


def formatar_linhas(d: dict[str, int]) -> str:
    return ", ".join(f"{nome}: {qtd}" for nome, qtd in d.items())


def delta_linhas(antes: dict[str, int], depois: dict[str, int]) -> list[str]:
    """Gera linhas descritivas com a variação de linhas por planilha."""
    linhas = []
    for nome in antes:
        a = antes[nome]
        d = depois.get(nome, 0)
        variacao = "N/D" if (a == -1 or d == -1) else str(d - a)
        linhas.append(f"  - {nome}: {a} -> {d}  (delta: {variacao})")
    return linhas


# ─── Processamento de cada arquivo ────────────────────────────────────────────

def processar_arquivo(caminho: str, excel_app, log: list[str]) -> dict:
    """Abre, atualiza, conta linhas, salva e fecha um único arquivo Excel."""
    nome = os.path.basename(caminho)
    separador = "=" * 70

    log += ["", separador, f"ARQUIVO : {nome}", f"Caminho : {caminho}", separador]

    resultado = {
        "arquivo": nome,
        "status": "FALHA",
        "erro": None,
        "tempo_atualizacao_s": None,
        "linhas_iniciais": {},
        "linhas_finais": {},
    }

    if not os.path.isfile(caminho):
        log.append(f"ERRO: arquivo nao encontrado: {caminho}")
        resultado["erro"] = "Arquivo não encontrado"
        return resultado

    workbook = None
    try:
        # ── Abertura ──────────────────────────────────────────────────────────
        log.append(f"[{_agora()}] Abrindo arquivo...")
        workbook = excel_app.Workbooks.Open(caminho, UpdateLinks=0, ReadOnly=False)
        log.append(f"[{_agora()}] Arquivo aberto com sucesso.")

        # ── Linhas antes ──────────────────────────────────────────────────────
        linhas_antes = contar_linhas_planilhas(workbook)
        resultado["linhas_iniciais"] = linhas_antes
        log.append(f"Linhas ANTES  -> {formatar_linhas(linhas_antes)}")

        # ── RefreshAll ────────────────────────────────────────────────────────
        log.append(f"[{_agora()}] Iniciando RefreshAll...")
        t0 = time.perf_counter()

        workbook.RefreshAll()

        try:
            excel_app.CalculateUntilAsyncQueriesDone()
        except AttributeError:
            log.append("Aviso: CalculateUntilAsyncQueriesDone indisponível. Aguardando 15 s.")
            time.sleep(15)

        elapsed = time.perf_counter() - t0
        resultado["tempo_atualizacao_s"] = round(elapsed, 2)
        log.append(f"[{_agora()}] RefreshAll concluído em {elapsed:.2f} s ({elapsed / 60:.2f} min).")

        if elapsed > TIMEOUT:
            log.append(f"ATENÇÃO: tempo de atualização excedeu o limite de {TIMEOUT} s.")

        # ── Linhas depois ─────────────────────────────────────────────────────
        linhas_depois = contar_linhas_planilhas(workbook)
        resultado["linhas_finais"] = linhas_depois
        log.append(f"Linhas DEPOIS -> {formatar_linhas(linhas_depois)}")

        log.append("Variação por planilha:")
        log.extend(delta_linhas(linhas_antes, linhas_depois))

        # ── Salvar ────────────────────────────────────────────────────────────
        log.append(f"[{_agora()}] Salvando arquivo...")
        workbook.Save()
        log.append(f"[{_agora()}] Arquivo salvo com sucesso.")

        resultado["status"] = "SUCESSO"

    except Exception as exc:
        log.append(f"[{_agora()}] ERRO durante o processamento:")
        log.append(traceback.format_exc())
        resultado["erro"] = str(exc)

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
                log.append(f"[{_agora()}] Arquivo fechado.")
            except Exception as exc_close:
                log.append(f"AVISO: erro ao fechar o arquivo: {exc_close}")

    return resultado


# ─── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    inicio = datetime.datetime.now()
    log: list[str] = [
        "=" * 70,
        "LOG DE ATUALIZAÇÃO DE RELATÓRIOS EXCEL",
        "=" * 70,
        f"Início      : {inicio.strftime('%d/%m/%Y %H:%M:%S')}",
        f"Arquivos    : {len(ARCHIVES)}",
    ]

    excel_app = None
    resultados: list[dict] = []

    try:
        log.append(f"[{_agora()}] Inicializando Excel em segundo plano...")
        excel_app = win32.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.AskToUpdateLinks = False

        for caminho in ARCHIVES:
            resultados.append(processar_arquivo(caminho, excel_app, log))

    except Exception:
        log.append("ERRO GERAL na execução do script:")
        log.append(traceback.format_exc())

    finally:
        if excel_app is not None:
            try:
                excel_app.Quit()
                log.append("")
                log.append(f"[{_agora()}] Aplicação Excel finalizada.")
            except Exception as exc_quit:
                log.append(f"AVISO: erro ao finalizar o Excel: {exc_quit}")

    fim = datetime.datetime.now()
    duracao = (fim - inicio).total_seconds()

    # ── Resumo ────────────────────────────────────────────────────────────────
    sucessos = sum(1 for r in resultados if r["status"] == "SUCESSO")
    falhas = len(resultados) - sucessos

    log += [
        "",
        "=" * 70,
        "RESUMO GERAL",
        "=" * 70,
        f"Início  : {inicio.strftime('%d/%m/%Y %H:%M:%S')}",
        f"Fim     : {fim.strftime('%d/%m/%Y %H:%M:%S')}",
        f"Duração : {duracao:.2f} s ({duracao / 60:.2f} min)",
        f"Sucesso : {sucessos}/{len(resultados)}",
        f"Falha   : {falhas}/{len(resultados)}",
        "",
    ]

    for r in resultados:
        log.append(f"  [{r['status']}] {r['arquivo']}")
        if r["status"] == "SUCESSO":
            log.append(f"    Tempo          : {r['tempo_atualizacao_s']} s")
            log.append(f"    Linhas iniciais: {formatar_linhas(r['linhas_iniciais'])}")
            log.append(f"    Linhas finais  : {formatar_linhas(r['linhas_finais'])}")
        else:
            log.append(f"    Erro: {r['erro']}")

    # ── Salvar log ────────────────────────────────────────────────────────────
    pasta_destino = PASTA_LOG if PASTA_LOG is not None else _SCRIPT_DIR  # CORRIGIDO: conflito de nome resolvido
    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)

    nome_log = f"log_atualizacao_{inicio.strftime('%Y%m%d_%H%M%S')}.txt"
    caminho_log = pasta_destino / nome_log

    conteudo = "\n".join(log)
    try:
        caminho_log.write_text(conteudo, encoding="utf-8")
        print(f"Log salvo em: {caminho_log}")
    except Exception as exc_log:
        print(f"ERRO ao salvar log: {exc_log}")

    print(conteudo)


if __name__ == "__main__":
    main()
