"""
MASTER EXCEL TOOLS - Web & API Server
Serves static frontend and provides high-speed, zero-RAM Excel/GST Cleaning API endpoints.
Engineered for Render (512MB RAM Limit) with strict disk-streamed processing.
"""

import os
import sys
import re
import uuid
import tempfile
from typing import Any, Optional, Dict, List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn

from cleaner_engine import clean_single_register_streaming

DIRECTORY = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("PORT", 8080))

app = FastAPI(
    title="MASTER EXCEL TOOLS",
    description="Enterprise Excel Merging & Zero-RAM GST Register Cleaning Engine",
    version="2.0.0"
)

# Enable CORS for cross-origin or local requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["*"]
)


@app.middleware("http")
async def add_no_cache_headers(request, call_next):
    """Prevent browsers from caching static HTML/JS/CSS assets to ensure zero stale versions."""
    response = await call_next(request)
    path = request.url.path
    if path == "/" or path.endswith((".html", ".js", ".css")):
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


def safe_remove_file(filepath: str):
    """Background task to cleanup temporary files without crashing."""
    try:
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
    except Exception as e:
        print(f"[CLEANUP WARNING] Could not remove {filepath}: {e}", file=sys.stderr)


@app.get("/api/health")
async def health_check():
    """Health check endpoint for Render / monitoring."""
    return {
        "status": "healthy",
        "service": "MASTER EXCEL TOOLS",
        "engine": "Zero-RAM Streaming Cleaner",
        "max_ram_limit": "512 MB Safe"
    }


@app.post("/api/clean-purge")
async def clean_purge_endpoint():
    """Immediately reclaims Python server memory via garbage collection."""
    import gc
    gc.collect()
    return {"status": "purged", "reclaimed": True}


def safe_header(val: Any, default: str = "") -> str:
    """Sanitize strings for HTTP headers (ASCII only, no newlines/returns)."""
    s = str(val if val is not None else default).strip().replace("\r", " ").replace("\n", " ")
    return s.encode("ascii", "replace").decode("ascii")


@app.post("/api/clean-register")
async def clean_register_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    register_type: str = Form("sale_erp"),
    output_format: str = Form("csv"),
    clean_filename: Optional[str] = Form(None)
):
    """
    Zero-RAM Streaming Upload & Cleaning Endpoint:
    1. Streams chunked upload directly to disk (RAM < 1MB).
    2. Runs SAX + SQLite streaming cleaner (Peak RAM < 35MB).
    3. Returns cleaned file with clean human-readable name & full statistics.
    4. Auto-cleans temp files in background.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="Missing filename in upload.")

    orig_ext = os.path.splitext(file.filename)[1].lower()
    if orig_ext not in (".xlsx", ".xlsm", ".csv"):
        raise HTTPException(status_code=400, detail=f"Unsupported format: {orig_ext}. Supported: .xlsx, .xlsm, .csv")

    temp_dir = tempfile.gettempdir()
    unique_id = uuid.uuid4().hex
    
    # Safe disk filename
    safe_input_name = re.sub(r'[^a-zA-Z0-9_.-]', '_', os.path.basename(file.filename or "input.xlsx"))
    input_path = os.path.join(temp_dir, f"in_{unique_id}_{safe_input_name}")

    target_ext = ".csv" if output_format.lower() == "csv" else ".xlsx"
    
    # Generate clean human-readable output filename
    if clean_filename and clean_filename.strip():
        base_clean = clean_filename.strip()
        if not base_clean.lower().endswith(target_ext):
            out_filename = f"{base_clean}{target_ext}"
        else:
            out_filename = base_clean
    else:
        clean_base = os.path.splitext(safe_input_name)[0]
        out_filename = f"Cleaned_{clean_base}{target_ext}"

    safe_out_name = re.sub(r'[^a-zA-Z0-9_.-]', '_', out_filename)
    output_path = os.path.join(temp_dir, f"out_{unique_id}_{safe_out_name}")

    try:
        # Step 1: High-Speed Chunked File Upload (1 MB chunks)
        with open(input_path, "wb") as f_out:
            while True:
                chunk = await file.read(1048576)
                if not chunk:
                    break
                f_out.write(chunk)

        # Step 2: Stream Clean File
        stats = clean_single_register_streaming(
            input_filepath=input_path,
            output_filepath=output_path,
            register_type=register_type,
            output_format=output_format
        )

        # Step 3: Schedule background file cleanup & Python memory reclamation
        background_tasks.add_task(safe_remove_file, input_path)
        background_tasks.add_task(safe_remove_file, output_path)
        import gc
        background_tasks.add_task(gc.collect)

        media_type = "text/csv" if target_ext == ".csv" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        headers = {
            "Access-Control-Expose-Headers": "*",
            "X-Stats-Input-File": safe_header(file.filename, "input.xlsx"),
            "X-Stats-Output-File": safe_header(out_filename, "Cleaned_Register.xlsx"),
            "X-Stats-Column-Detected": safe_header(stats.get("column_detected"), "Auto Detected"),
            "X-Stats-Column-Index": safe_header(stats.get("column_index"), "0"),
            "X-Stats-Header-Row": safe_header(stats.get("header_row_index"), "0"),
            "X-Stats-Original-Rows": safe_header(stats.get("original_row_count"), "0"),
            "X-Stats-Retained-Rows": safe_header(stats.get("retained_row_count"), "0"),
            "X-Stats-Excluded-Rows": safe_header(stats.get("excluded_row_count"), "0"),
            "X-Stats-Peak-RAM-MB": safe_header(stats.get("memory_peak_mb"), "0.0"),
            "X-Stats-Duration-Sec": safe_header(stats.get("duration_seconds"), "0.0")
        }

        return FileResponse(
            path=output_path,
            filename=safe_header(out_filename, "Cleaned_Register.xlsx"),
            media_type=media_type,
            headers=headers
        )

    except Exception as e:
        safe_remove_file(input_path)
        safe_remove_file(output_path)
        print(f"[CLEANING ERROR] {e}", file=sys.stderr)
        raise HTTPException(status_code=500, detail=f"Cleaning failed: {str(e)}")


# Mount Static Frontend (index.html, styles.css, app.js)
app.mount("/", StaticFiles(directory=DIRECTORY, html=True), name="static")

if __name__ == "__main__":
    print(f"MASTER EXCEL TOOLS & CLEANING API running at http://localhost:{PORT}/", flush=True)
    uvicorn.run("server:app", host="0.0.0.0", port=PORT, reload=False, log_level="info")
