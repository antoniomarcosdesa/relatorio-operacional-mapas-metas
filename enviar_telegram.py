# ============================================================
# ENVIAR_TELEGRAM.PY
#
# Envia o PDF do relatório para um chat/grupo do Telegram.
# Credenciais lidas de telegram.env (na raiz da pasta do cliente):
#   TELEGRAM_TOKEN=...   TELEGRAM_CHAT_ID=...
# (aceita também o formato  TELEGRAM_TOKEN = '...';  )
# Só usa a biblioteca padrão (sem instalar nada).
# ============================================================

import json
import os
import time
import uuid
import urllib.request
import urllib.error
from pathlib import Path

API_BASE = os.environ.get("TELEGRAM_API_BASE", "https://api.telegram.org")


def carregar_config(raiz: Path):
    arq = Path(raiz) / "telegram.env"
    cfg = {}
    if arq.exists():
        for linha in arq.read_text(encoding="utf-8").splitlines():
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            k, v = linha.split("=", 1)
            cfg[k.strip()] = v.strip().rstrip(";").strip().strip("'\"")
    # variáveis de ambiente têm prioridade, se existirem
    for k in ("TELEGRAM_TOKEN", "TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID"):
        if os.environ.get(k):
            cfg[k] = os.environ[k]
    return cfg.get("TELEGRAM_TOKEN") or cfg.get("TELEGRAM_BOT_TOKEN"), cfg.get("TELEGRAM_CHAT_ID")


def _enviar_urllib(caminho_pdf, legenda, token, chat_id):
    limite = 'bnd' + uuid.uuid4().hex
    corpo = b""
    for nome, valor in (("chat_id", chat_id), ("caption", legenda)):
        corpo += (f'--{limite}\r\nContent-Disposition: form-data; name="{nome}"\r\n\r\n{valor}\r\n').encode("utf-8")
    corpo += (f'--{limite}\r\nContent-Disposition: form-data; name="document"; filename="{caminho_pdf.name}"\r\n'
              f'Content-Type: application/pdf\r\n\r\n').encode("utf-8")
    corpo += caminho_pdf.read_bytes() + f"\r\n--{limite}--\r\n".encode("utf-8")
    req = urllib.request.Request(
        f"{API_BASE}/bot{token}/sendDocument", data=corpo,
        headers={"Content-Type": f"multipart/form-data; boundary={limite}"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return r.status, json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except Exception:
            return e.code, {}


def _enviar_requests(caminho_pdf, legenda, token, chat_id):
    """Mesmo método do enviar_telegram.py do projeto SGPA (requests)."""
    import requests
    with open(caminho_pdf, "rb") as f:
        r = requests.post(f"{API_BASE}/bot{token}/sendDocument",
                          data={"chat_id": chat_id, "caption": legenda},
                          files={"document": (caminho_pdf.name, f, "application/pdf")},
                          timeout=(15, 180))
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, {}


def enviar_pdf(caminho_pdf, legenda, raiz, tentativas=3):
    """Devolve (ok, mensagem). Nunca levanta exceção. Usa requests (como o projeto
    SGPA); sem requests instalado, usa urllib. Tenta até 3 vezes."""
    token, chat_id = carregar_config(raiz)
    if not token or not chat_id:
        return False, "TELEGRAM_TOKEN / TELEGRAM_CHAT_ID não configurados em telegram.env."
    caminho_pdf = Path(caminho_pdf)
    try:
        import requests  # noqa: F401
        enviar = _enviar_requests
    except ImportError:
        enviar = _enviar_urllib
    ultimo = ""
    for t in range(1, tentativas + 1):
        try:
            codigo, resp = enviar(caminho_pdf, legenda, token, chat_id)
            if resp.get("ok"):
                return True, "Relatório enviado ao Telegram."
            ultimo = f"HTTP {codigo}: {resp.get('description', resp) or 'sem detalhe'}"
            if codigo in (400, 401, 403, 404):  # erro de configuração: repetir não resolve
                break
        except Exception as e:
            ultimo = f"{type(e).__name__}: {e}"
        if t < tentativas:
            print(f"  Falha no envio ({ultimo}). Nova tentativa {t + 1}/{tentativas}...")
            time.sleep(3 * t)
    dica = ""
    if "HTTP 401" in ultimo:
        dica = " -> token inválido/revogado."
    elif "chat not found" in ultimo.lower() or "HTTP 403" in ultimo or "kicked" in ultimo.lower():
        dica = " -> o bot não está no grupo (ou foi removido) ou o chat id está errado."
    elif "upgraded to a supergroup" in ultimo.lower():
        dica = " -> o grupo virou supergrupo; o chat id novo começa com -100."
    elif "SSL" in ultimo or "certificate" in ultimo.lower():
        dica = " -> problema de certificado/proxy da rede."
    return False, f"Não consegui enviar ao Telegram: {ultimo}{dica}"


if __name__ == "__main__":
    # Teste manual:  python scripts/enviar_telegram.py caminho\\do\\relatorio.pdf
    import sys
    if len(sys.argv) < 2:
        print("Uso: python scripts/enviar_telegram.py <arquivo.pdf>")
        sys.exit(1)
    ok, msg = enviar_pdf(sys.argv[1], "Teste de envio — relatório", Path(__file__).resolve().parent.parent)
    print(("[OK] " if ok else "[ERRO] ") + msg)
    sys.exit(0 if ok else 1)
