"""
Financial Model API Server
===========================
Deploy this on your VPS alongside financial_model_v5.py.
Exposes a single POST endpoint that n8n (or your frontend) calls.

Usage:
  pip install fastapi uvicorn anthropic yfinance beautifulsoup4 requests openpyxl pandas
  ANTHROPIC_API_KEY="sk-ant-..." python3 financial_model_server.py

The server runs on port 8500 by default.

Deployment setup (one-time, as root):
  sudo mkdir -p /var/lib/financial_models
  sudo chown <deploy-user> /var/lib/financial_models
  # Optional: prevent systemd-tmpfiles from clearing it
  echo "x /var/lib/financial_models" | sudo tee /etc/tmpfiles.d/financial_models.conf
"""

import os
import sys
import uuid
import time
import traceback
from fastapi import FastAPI, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import requests
import uvicorn

# ── Make sure financial_model_v3 is importable from the same directory ──
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from financial_model_v5 import __version__ as FM_VERSION

app = FastAPI(title="Financial Model Generator", version="1.0")

# Output directory for generated models
OUTPUT_DIR = os.environ.get("MODEL_OUTPUT_DIR", "/var/lib/financial_models")
try:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
except PermissionError as e:
    raise RuntimeError(
        f"Cannot create {OUTPUT_DIR}. Run: sudo mkdir -p {OUTPUT_DIR} && sudo chown $(whoami) {OUTPUT_DIR}"
    ) from e
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
MODEL_STORAGE_BUCKET = os.environ.get("SUPABASE_FINANCIAL_MODEL_BUCKET", "research-reports-html")

# Track job status for async mode
jobs: dict[str, dict] = {}


# ========================
# Request / Response Models
# ========================

class GenerateRequest(BaseModel):
    nse_symbol: str
    company_name: str
    sector: str
    folder_id: str | None = None  # Optional — for reference only


class GenerateResponse(BaseModel):
    status: str  # "success" | "error"
    file_name: str | None = None
    file_path: str | None = None
    storage_path: str | None = None
    storage_url: str | None = None
    json_storage_path: str | None = None
    json_storage_url: str | None = None
    message: str | None = None
    duration_seconds: int | None = None


class JobStatus(BaseModel):
    job_id: str
    status: str  # "processing" | "completed" | "failed"
    file_name: str | None = None
    file_path: str | None = None
    storage_path: str | None = None
    storage_url: str | None = None
    json_storage_path: str | None = None
    json_storage_url: str | None = None
    message: str | None = None
    duration_seconds: int | None = None


class StorageMirrorResponse(BaseModel):
    status: str  # "success" | "error"
    file_name: str | None = None
    file_path: str | None = None
    storage_path: str | None = None
    storage_url: str | None = None
    json_storage_path: str | None = None
    json_storage_url: str | None = None
    message: str | None = None


# ========================
# Health Check
# ========================

@app.get("/health")
def health():
    return {
        "status": "ok",
        "model_version": FM_VERSION,
        "anthropic_key_set": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "screener_creds_set": bool(os.environ.get("SCREENER_USERNAME") and os.environ.get("SCREENER_PASSWORD")),
        "output_dir": OUTPUT_DIR,
    }


def _storage_public_url(path: str) -> str:
    return f"{SUPABASE_URL}/storage/v1/object/public/{MODEL_STORAGE_BUCKET}/{path}"


def upload_model_to_supabase(file_path: str, ticker: str) -> dict[str, str | None]:
    """Upload the v5 xlsx and (when present) the .json sidecar produced by
    financial_model_v5.generate_financial_model. The html-report pipeline
    prefers the JSON sidecar because the xlsx is formula-only and openpyxl
    can't read computed cell values without Excel having recalculated them.
    """
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise RuntimeError("SUPABASE_URL / SUPABASE_SERVICE_KEY not configured")

    ticker_u = ticker.upper()
    path = f"financial-models/{ticker_u}/{ticker_u}_model.xlsx"
    url = f"{SUPABASE_URL}/storage/v1/object/{MODEL_STORAGE_BUCKET}/{path}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "apikey": SUPABASE_SERVICE_KEY,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "x-upsert": "true",
        "Cache-Control": "no-cache",
    }

    with open(file_path, "rb") as f:
        resp = requests.put(url, data=f, headers=headers, timeout=120)

    if resp.status_code >= 400:
        raise RuntimeError(f"Supabase storage upload failed: {resp.status_code} {resp.text[:400]}")

    # Also upload the JSON sidecar if it exists alongside the xlsx
    json_local = os.path.splitext(file_path)[0] + ".json"
    json_path = None
    json_public_url = None
    if os.path.exists(json_local):
        json_path = f"financial-models/{ticker_u}/{ticker_u}_model.json"
        json_url = f"{SUPABASE_URL}/storage/v1/object/{MODEL_STORAGE_BUCKET}/{json_path}"
        json_headers = {
            **headers,
            "Content-Type": "application/json; charset=utf-8",
        }
        try:
            with open(json_local, "rb") as f:
                jr = requests.put(json_url, data=f, headers=json_headers, timeout=60)
            if jr.status_code >= 400:
                raise RuntimeError(f"JSON sidecar upload failed: {jr.status_code} {jr.text[:200]}")
            print(f"  Uploaded JSON sidecar -> {json_path}")
            json_public_url = _storage_public_url(json_path)
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(f"JSON sidecar upload failed ({type(e).__name__}: {e})") from e
    else:
        print(f"  NOTE: no JSON sidecar at {json_local}; html pipeline will fall back to CSV/xlsx")

    return {
        "storage_path": path,
        "storage_url": _storage_public_url(path),
        "json_storage_path": json_path,
        "json_storage_url": json_public_url,
    }


# ========================
# Synchronous Generation (n8n calls this with long timeout)
# ========================

@app.post("/generate", response_model=GenerateResponse)
def generate_sync(req: GenerateRequest):
    """
    Synchronous endpoint — blocks until the model is generated.
    n8n should call this with a 15-minute timeout.
    Returns the file path so n8n can read and upload it.
    """
    ticker = req.nse_symbol.strip().upper()
    start = time.time()

    print(f"[generate] Received: nse_symbol='{req.nse_symbol}', ticker='{ticker}', company='{req.company_name}', sector='{req.sector}'")

    if not ticker:
        return GenerateResponse(
            status="error",
            message="nse_symbol is empty — check the webhook payload",
            duration_seconds=0,
        )

    try:
        # v5: formula-based, year-agnostic, cost-tracked
        from financial_model_v5 import generate_financial_model

        out_dir = os.path.join(OUTPUT_DIR, ticker)
        os.makedirs(out_dir, exist_ok=True)

        result = generate_financial_model(
            nse_code=ticker,
            company_name=req.company_name,
            sector=req.sector,
            output_dir=out_dir,
        )

        file_path = result["file_path"]
        json_src = result.get("json_path")

        # Copy xlsx + JSON sidecar into the canonical OUTPUT_DIR slot so
        # upload_model_to_supabase can find both side-by-side.
        final_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")
        final_json = os.path.join(OUTPUT_DIR, f"{ticker}_model.json")
        if file_path != final_path:
            import shutil
            shutil.copy2(file_path, final_path)
            file_path = final_path
        if json_src and os.path.exists(json_src) and json_src != final_json:
            import shutil
            shutil.copy2(json_src, final_json)

        storage_result = upload_model_to_supabase(file_path, ticker)

        return GenerateResponse(
            status="success",
            file_name=f"{ticker}_model.xlsx",
            file_path=file_path,
            storage_path=storage_result["storage_path"],
            storage_url=storage_result["storage_url"],
            json_storage_path=storage_result["json_storage_path"],
            json_storage_url=storage_result["json_storage_url"],
            duration_seconds=int(time.time() - start),
            message=(
                f"rating={result.get('rating')} target={result.get('target_price')} "
                f"upside={result.get('upside_pct')}% cost=${result['cost_summary']['cost_usd']}"
            ),
        )

    except Exception as e:
        traceback.print_exc()
        return GenerateResponse(
            status="error",
            message=str(e),
            duration_seconds=int(time.time() - start),
        )


# ========================
# Async Generation (if you want non-blocking)
# ========================

@app.post("/generate-async")
def generate_async(req: GenerateRequest, background_tasks: BackgroundTasks):
    """
    Async endpoint — returns immediately with a job_id.
    Poll /job/{job_id} to check status.
    """
    job_id = str(uuid.uuid4())[:8]
    jobs[job_id] = {"status": "processing", "started_at": time.time()}

    background_tasks.add_task(_run_generation, job_id, req)

    return {"job_id": job_id, "status": "processing", "message": "Generation started"}


@app.get("/job/{job_id}", response_model=JobStatus)
def get_job_status(job_id: str):
    """Check the status of an async generation job."""
    if job_id not in jobs:
        return JSONResponse(status_code=404, content={"error": "Job not found"})

    job = jobs[job_id]
    return JobStatus(
        job_id=job_id,
        status=job["status"],
        file_name=job.get("file_name"),
        file_path=job.get("file_path"),
        storage_path=job.get("storage_path"),
        storage_url=job.get("storage_url"),
        json_storage_path=job.get("json_storage_path"),
        json_storage_url=job.get("json_storage_url"),
        message=job.get("message"),
        duration_seconds=job.get("duration_seconds"),
    )


def _run_generation(job_id: str, req: GenerateRequest):
    """Background task for async generation."""
    ticker = req.nse_symbol.upper()
    start = time.time()

    try:
        from financial_model_v5 import generate_financial_model

        result = generate_financial_model(
            nse_code=ticker,
            company_name=req.company_name,
            sector=req.sector,
            output_dir=os.path.join(OUTPUT_DIR, ticker),
        )
        file_path = result["file_path"]
        json_src = result.get("json_path")

        final_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")
        final_json = os.path.join(OUTPUT_DIR, f"{ticker}_model.json")
        if file_path != final_path:
            import shutil
            shutil.copy2(file_path, final_path)
            file_path = final_path
        if json_src and os.path.exists(json_src) and json_src != final_json:
            import shutil
            shutil.copy2(json_src, final_json)

        storage_result = upload_model_to_supabase(file_path, ticker)

        jobs[job_id] = {
            "status": "completed",
            "file_name": f"{ticker}_model.xlsx",
            "file_path": file_path,
            "storage_path": storage_result["storage_path"],
            "storage_url": storage_result["storage_url"],
            "json_storage_path": storage_result["json_storage_path"],
            "json_storage_url": storage_result["json_storage_url"],
            "duration_seconds": int(time.time() - start),
            "rating": result.get("rating"),
            "target_price": result.get("target_price"),
            "upside_pct": result.get("upside_pct"),
            "cost_usd": result["cost_summary"]["cost_usd"],
        }

    except Exception as e:
        traceback.print_exc()
        jobs[job_id] = {
            "status": "failed",
            "message": str(e),
            "duration_seconds": int(time.time() - start),
        }


# ========================
# Download the generated file
# ========================

@app.get("/download/{ticker}")
def download_model(ticker: str):
    """Download the generated Excel file."""
    ticker = ticker.upper()
    file_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")

    if not os.path.exists(file_path):
        return JSONResponse(status_code=404, content={"error": f"No model found for {ticker}"})

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{ticker}_model.xlsx",
    )


@app.post("/storage/{ticker}", response_model=StorageMirrorResponse)
def mirror_model_to_storage(ticker: str):
    ticker = ticker.upper()
    file_path = os.path.join(OUTPUT_DIR, f"{ticker}_model.xlsx")

    if not os.path.exists(file_path):
        return StorageMirrorResponse(
            status="error",
            message=f"No model found for {ticker}",
        )

    try:
        storage_result = upload_model_to_supabase(file_path, ticker)
        return StorageMirrorResponse(
            status="success",
            file_name=f"{ticker}_model.xlsx",
            file_path=file_path,
            storage_path=storage_result["storage_path"],
            storage_url=storage_result["storage_url"],
            json_storage_path=storage_result["json_storage_path"],
            json_storage_url=storage_result["json_storage_url"],
        )
    except Exception as e:
        traceback.print_exc()
        return StorageMirrorResponse(
            status="error",
            file_name=f"{ticker}_model.xlsx",
            file_path=file_path,
            message=str(e),
        )


# ========================
# Run Server
# ========================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8500))
    print(f"\n🚀 Financial Model Server starting on port {port}")
    print(f"📁 Output directory: {OUTPUT_DIR}")
    print(f"🔑 Anthropic key: {'✓ Set' if os.environ.get('ANTHROPIC_API_KEY') else '✗ MISSING'}\n")

    uvicorn.run(app, host="0.0.0.0", port=port)
