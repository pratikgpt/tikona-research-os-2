"""
PPTX Generation API Server
==========================
Single-purpose FastAPI service that turns approved research reports into
.pptx (and optional .pdf) using the local `reportgen` library at ../../pptx.

Required env vars:
  SUPABASE_URL
  SUPABASE_SERVICE_KEY
  REPORTGEN_OPENROUTER_API_KEY    # used by reportgen planner when use_mock=False
  PORT                             # optional, default 8501

Run:
  pip install -e ../../pptx           # one-time, install the local library
  pip install -r requirements.txt
  python main.py
"""
from __future__ import annotations

import logging
import os
import shutil
import sys
from pathlib import Path
import time
import traceback

from dotenv import load_dotenv

SERVICE_DIR = Path(__file__).resolve().parent
REPO_ROOT = SERVICE_DIR.parent.parent

# Load env files explicitly so the service behaves the same whether it is
# launched from the repo root, scripts/ppt_service, or by another process.
load_dotenv(REPO_ROOT / ".env")
load_dotenv(SERVICE_DIR / ".env", override=False)
load_dotenv()

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import uvicorn
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(REPO_ROOT / "PPT-generator"))

from pptx_generator import generate_pptx_for_report, preview_ppt_placeholders

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("ppt_service")


app = FastAPI(title="PPTX Generation Service", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://localhost:3000",
        "http://72.61.226.16",
        "http://72.61.226.16:8501",
    ],
    allow_origin_regex=r"https://.*\.vercel\.app",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── request / response models ──────────────────────────────────────────────


class GeneratePptxRequest(BaseModel):
    reportId: str
    sessionId: str
    useMock: bool = False


class PreviewPlaceholdersRequest(BaseModel):
    reportId: str
    sessionId: str


class GeneratePptxResponse(BaseModel):
    status: str
    message: str
    pptx_file_url: str
    pptx_file_path: str
    pptx_pdf_file_url: str | None = None
    pptx_pdf_file_path: str | None = None
    duration_seconds: float
    warnings: list[str] = []


# ── routes ─────────────────────────────────────────────────────────────────


@app.get("/health")
def health():
    reportgen_ok = False
    try:
        import reportgen.orchestration.pipeline  # noqa: F401

        reportgen_ok = True
    except Exception as exc:
        logger.warning("reportgen import failed: %s", exc)

    return {
        "status": "ok",
        "reportgen": reportgen_ok,
        "libreoffice": shutil.which("soffice") is not None,
        "supabase": bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SERVICE_KEY")),
        "openrouter_key": bool(os.environ.get("REPORTGEN_OPENROUTER_API_KEY")),
    }


@app.post("/preview-placeholders")
def preview_placeholders(req: PreviewPlaceholdersRequest):
    logger.info("POST /preview-placeholders report=%s session=%s", req.reportId, req.sessionId)
    try:
        result = preview_ppt_placeholders(req.reportId, req.sessionId)
        return result
    except Exception as exc:
        logger.error("preview_placeholders failed: %s\n%s", exc, traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": str(exc)[:500]},
        )


@app.post("/generate-pptx", response_model=GeneratePptxResponse)
def generate_pptx(req: GeneratePptxRequest):
    t0 = time.time()
    logger.info("POST /generate-pptx report=%s session=%s use_mock=%s", req.reportId, req.sessionId, req.useMock)
    try:
        result = generate_pptx_for_report(req.reportId, req.sessionId, use_mock=req.useMock)
        return GeneratePptxResponse(**result)
    except Exception as exc:
        duration = round(time.time() - t0, 2)
        logger.error("generate_pptx failed: %s\n%s", exc, traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "message": str(exc)[:500],
                "duration_seconds": duration,
            },
        )


# ── entrypoint ─────────────────────────────────────────────────────────────


def _print_banner(port: int) -> None:
    try:
        import reportgen  # noqa: F401

        rg_version = getattr(reportgen, "__version__", "editable")
    except Exception:
        rg_version = "NOT INSTALLED"

    has_or = bool(os.environ.get("REPORTGEN_OPENROUTER_API_KEY"))
    has_sb = bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SERVICE_KEY"))
    libre = shutil.which("soffice") is not None

    bar = "=" * 60
    print(bar)
    print(f" PPTX Generation Service  ->  http://0.0.0.0:{port}")
    print(f"  reportgen        : {rg_version}")
    print(f"  openrouter key   : {'yes' if has_or else 'MISSING (use_mock=true required)'}")
    print(f"  supabase env     : {'yes' if has_sb else 'MISSING'}")
    print(f"  libreoffice      : {'yes' if libre else 'no (PDF will be skipped)'}")
    print(f"  bucket           : research-reports-pptx")
    print(bar)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8501))
    _print_banner(port)
    uvicorn.run(app, host="0.0.0.0", port=port)
