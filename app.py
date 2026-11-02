#!/usr/bin/env python3
"""
Servidor FastAPI para o Conversor Universal → Markdown.

Endpoints:
  GET  /                -> página HTML com a interface (upload de arquivo ou URL)
  POST /api/convert     -> converte arquivo enviado (multipart) OU URL remota
  GET  /static/{path}   -> assets estáticos (CSS, JS)

Como executar:
  pip install -r requirements.txt
  uvicorn app:app --reload --host 0.0.0.0 --port 8000
"""

from __future__ import annotations

import logging
import re
import tempfile
from pathlib import Path
from typing import Optional

import httpx
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from conversor_unificado import (
    ConversionStats,
    convert,
    convert_bytes,
    convert_html_text,
    detect_format,
    remove_striked_paragraphs,
)

LOGGER = logging.getLogger("conversor_md_api")
logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(name)s | %(message)s")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

# Limite de tamanho para upload e download (50 MB) — síncrono, evita OOM
MAX_BYTES = 50 * 1024 * 1024
# Timeout para download de URL (30 s)
URL_TIMEOUT = 30.0
# User-Agent para evitar bloqueios em portais governamentais
USER_AGENT = "ConversorMD/1.0 (+https://github.com/conversor-md)"

# Hosts permitidos para fetch de URL (allowlist de portais jurídicos brasileiros comuns)
# Se vazio, aceita qualquer URL. Ajuste conforme sua necessidade de segurança.
ALLOWED_HOSTS: list[str] = []  # ex: ["planalto.gov.br", "in.gov.br", "stf.jus.br"]

app = FastAPI(
    title="Conversor Universal → Markdown",
    description="Converte PDF, HTML (com remoção de revogados), Office, imagens e texto para Markdown.",
    version="1.0.0",
)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


# =============================================================================
# Utils
# =============================================================================
def _sanitize_filename(name: str) -> str:
    """Sanitiza nome de arquivo para uso seguro em Content-Disposition."""
    # Remove caracteres problemáticos, mantém unicode
    name = re.sub(r'[\x00-\x1f<>:"/\\|?*]', "_", name).strip().strip(".")
    return name or "documento"


def _derive_md_filename(source_name: str) -> str:
    """Gera nome do .md a partir do nome de origem."""
    stem = Path(source_name).stem or "documento"
    return _sanitize_filename(stem) + ".md"


def _validate_url(url: str) -> None:
    """Validações básicas de URL: esquema e (se configurado) host allowlist."""
    url = url.strip()
    if not url:
        raise HTTPException(status_code=400, detail="URL vazia.")
    if not re.match(r"^https?://", url, re.IGNORECASE):
        raise HTTPException(status_code=400, detail="URL deve começar com http:// ou https://")
    if ALLOWED_HOSTS:
        from urllib.parse import urlparse
        host = urlparse(url).hostname or ""
        if not any(host == h or host.endswith("." + h) for h in ALLOWED_HOSTS):
            raise HTTPException(
                status_code=403,
                detail=f"Host '{host}' não está na lista de permitidos. Contate o administrador.",
            )


async def _fetch_url(url: str) -> tuple[bytes, str, str]:
    """
    Baixa conteúdo da URL e retorna (content, content_type, filename_suggestion).
    Lança HTTPException em caso de falha.
    """
    try:
        async with httpx.AsyncClient(
            follow_redirects=True,
            timeout=httpx.Timeout(URL_TIMEOUT),
            headers={"User-Agent": USER_AGENT},
        ) as client:
            resp = await client.get(url)
            resp.raise_for_status()
    except httpx.TimeoutException:
        raise HTTPException(status_code=504, detail=f"Timeout ao baixar a URL (limite de {URL_TIMEOUT:.0f}s).")
    except httpx.HTTPError as exc:
        raise HTTPException(status_code=502, detail=f"Falha ao baixar a URL: {exc}")

    content = resp.content
    if len(content) > MAX_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Conteúdo da URL excede o limite de {MAX_BYTES // (1024 * 1024)} MB.",
        )

    content_type = resp.headers.get("content-type", "")
    # Tenta extrair filename do Content-Disposition ou do path da URL
    filename_suggestion = ""
    cd = resp.headers.get("content-disposition", "")
    if cd:
        m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd, re.IGNORECASE)
        if m:
            filename_suggestion = m.group(1)
    if not filename_suggestion:
        from urllib.parse import urlparse, unquote
        path = urlparse(url).path
        if path:
            filename_suggestion = unquote(Path(path).name) or "documento"

    return content, content_type, filename_suggestion


def _build_stats_dict(stats: ConversionStats) -> dict:
    """Converte ConversionStats em dict para JSON response."""
    return {
        "format": stats.format,
        "pages": stats.pages,
        "tables": stats.tables,
        "links": stats.links,
        "headings": stats.headings,
        "list_items": stats.list_items,
        "image_only_pages": stats.image_only_pages,
        "page_markers": stats.page_markers,
        "unresolved_link_urls": stats.unresolved_link_urls,
        "warnings": stats.warnings or [],
    }


# =============================================================================
# Rotas
# =============================================================================
@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    """Serve a página HTML principal."""
    index_path = TEMPLATES_DIR / "index.html"
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.post("/api/convert")
async def api_convert(
    file: Optional[UploadFile] = File(None),
    url: Optional[str] = Form(None),
    include_page_markers: bool = Form(True),
):
    """
    Converte um arquivo enviado (multipart/form-data) OU uma URL remota.

    - Se `file` for enviado: usa o conteúdo + nome do arquivo.
    - Se `url` for enviada: baixa o conteúdo (detecção automática por Content-Type).
    - `include_page_markers`: se False, omite os comentários de página no MD (PDF).

    Retorna JSON com:
      - markdown (string)
      - filename (sugestão de nome do .md)
      - stats (métricas da conversão)
    """
    if not file and not url:
        raise HTTPException(status_code=400, detail="Envie um arquivo ou uma URL.")

    if file and url:
        raise HTTPException(status_code=400, detail="Envie apenas um: arquivo OU URL (não ambos).")

    if file:
        # === Upload direto ===
        content = await file.read()
        if len(content) > MAX_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"Arquivo excede o limite de {MAX_BYTES // (1024 * 1024)} MB.",
            )
        filename = file.filename or "documento"
        content_type = file.content_type or ""
        LOGGER.info("Upload recebido: %s (%s, %d bytes)", filename, content_type, len(content))
    else:
        # === URL remota ===
        assert url is not None
        _validate_url(url)
        content, content_type, filename = await _fetch_url(url)
        LOGGER.info("URL baixada: %s -> %s (%s, %d bytes)", url, filename, content_type, len(content))

    # === Conversão ===
    try:
        markdown, stats = convert_bytes(
            content=content,
            filename=filename,
            content_type=content_type,
            include_page_markers=include_page_markers,
        )
    except RuntimeError as exc:
        # Erro de dependência ausente (markitdown, pytesseract, etc.)
        LOGGER.error("Dependência ausente: %s", exc)
        raise HTTPException(status_code=500, detail=str(exc))
    except Exception as exc:
        LOGGER.exception("Falha na conversão")
        raise HTTPException(status_code=422, detail=f"Falha ao converter: {exc}")

    md_filename = _derive_md_filename(filename)

    return JSONResponse({
        "markdown": markdown,
        "filename": md_filename,
        "source_filename": filename,
        "source_content_type": content_type,
        "stats": _build_stats_dict(stats),
    })


@app.post("/api/convert/download")
async def api_convert_download(
    file: Optional[UploadFile] = File(None),
    url: Optional[str] = Form(None),
    include_page_markers: bool = Form(True),
):
    """
    Igual a /api/convert, mas retorna o .md diretamente como download
    (Content-Disposition: attachment) — útil para chamadas programáticas.
    """
    if not file and not url:
        raise HTTPException(status_code=400, detail="Envie um arquivo ou uma URL.")
    if file and url:
        raise HTTPException(status_code=400, detail="Envie apenas um: arquivo OU URL (não ambos).")

    if file:
        content = await file.read()
        if len(content) > MAX_BYTES:
            raise HTTPException(status_code=413, detail=f"Arquivo excede o limite de {MAX_BYTES // (1024 * 1024)} MB.")
        filename = file.filename or "documento"
        content_type = file.content_type or ""
    else:
        assert url is not None
        _validate_url(url)
        content, content_type, filename = await _fetch_url(url)

    try:
        markdown, _stats = convert_bytes(
            content=content,
            filename=filename,
            content_type=content_type,
            include_page_markers=include_page_markers,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))
    except Exception as exc:
        LOGGER.exception("Falha na conversão")
        raise HTTPException(status_code=422, detail=f"Falha ao converter: {exc}")

    md_filename = _derive_md_filename(filename)
    safe = _sanitize_filename(md_filename)

    return Response(
        content=markdown.encode("utf-8"),
        media_type="text/markdown; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{safe}"; filename*=UTF-8\'\'{safe}'},
    )


@app.get("/api/health")
async def health() -> dict:
    return {"status": "ok", "service": "conversor-md", "version": "1.0.0"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
