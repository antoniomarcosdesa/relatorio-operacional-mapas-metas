# Executa o pipeline (processar + PDF) numa pasta temporária isolada por execução e,
# opcionalmente, envia o PDF ao Telegram. Usado pelo app.py (site).
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
PASTA_SCRIPTS = RAIZ / "scripts"
PASTA_CLIENTES = RAIZ / "clientes"


def listar_clientes():
    return sorted(p.name for p in PASTA_CLIENTES.iterdir() if p.is_dir()) if PASTA_CLIENTES.exists() else []


def _rodar(script, args, cwd, log):
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "MPLBACKEND": "Agg"}
    proc = subprocess.Popen([sys.executable, str(PASTA_SCRIPTS / script), *args], cwd=cwd, env=env,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                            encoding="utf-8", errors="replace")
    for linha in proc.stdout:
        log(linha.rstrip())
    proc.wait()
    return proc.returncode == 0


def gerar_relatorio(csv_bytes: bytes, cliente: str, enviar_telegram: bool, log=print):
    """Devolve (ok, mensagem, caminho_pdf_ou_None, bytes_pdf_ou_None)."""
    with tempfile.TemporaryDirectory(prefix="rel_") as tmp:
        tmp = Path(tmp)
        for d in ("dados_brutos", "resultados", "relatorios", "cliente"):
            (tmp / d).mkdir()
        origem = PASTA_CLIENTES / cliente
        for arq in [a for e in ("*.xlsx", "*.kml", "*.geojson", "*.json", "*.yms", "*.ymx") for a in origem.glob(e)]:
            shutil.copy(arq, tmp / "cliente" / arq.name)
        csv = tmp / "dados_brutos" / "mapa_analitico.csv"
        csv.write_bytes(csv_bytes)

        log("Etapa 1/2 — processando indicadores...")
        if not _rodar("processar_mapa_analitico.py", [str(csv), "resultados"], tmp, log):
            return False, "Falhou ao processar os dados (veja o log).", None
        log("Etapa 2/2 — gerando mapas e PDF...")
        if not _rodar("gerar_relatorio_qualidade.py", [str(csv)], tmp, log):
            return False, "Falhou ao gerar o PDF (veja o log).", None

        pdfs = sorted((tmp / "relatorios").glob("Relatorio_Qualidade_Operacional_*.pdf"))
        if not pdfs:
            return False, "PDF não foi gerado.", None
        pdf_bytes = pdfs[-1].read_bytes()
        nome_pdf = pdfs[-1].name

        msg = "Relatório gerado."
        if enviar_telegram:
            sys.path.insert(0, str(PASTA_SCRIPTS))
            from enviar_telegram import enviar_pdf
            ok, m = enviar_pdf(pdfs[-1], f"Relatório de Qualidade Operacional — {cliente}", origem)
            msg += " " + m
            if not ok:
                return False, msg, (nome_pdf, pdf_bytes)
        return True, msg, (nome_pdf, pdf_bytes)
