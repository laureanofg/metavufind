import asyncio
import logging
import os
from pathlib import Path

from fastapi import FastAPI, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from scrapers.engine import load_config, scrape_all_targets, scrape_target
from utils.excel import generate_excel

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("app")

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config" / "targets.yaml"

config = load_config(str(CONFIG_PATH))

app = FastAPI(title="MetaVufindScraping", version="1.0.0")

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

SESSION_STORE: dict[str, dict] = {}


@app.get("/", response_class=HTMLResponse)
async def index(request: Request, query: str = ""):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "query": query,
            "targets": config.targets,
        },
    )


@app.post("/api/search", response_class=HTMLResponse)
async def search_results(
    request: Request,
    query: str = Query(..., min_length=1),
    max_results: int = Query(50, ge=10, le=200),
):
    config.max_results_per_target = max_results

    return templates.TemplateResponse(
        "partials/results_container.html",
        {
            "request": request,
            "query": query,
            "max_results": max_results,
            "targets": config.targets,
        },
    )


@app.get("/api/search/{target_id}", response_class=HTMLResponse)
async def search_target(
    request: Request,
    target_id: str,
    query: str = Query(..., min_length=1),
    max_results: int = Query(50, ge=10, le=200),
):
    target = next((t for t in config.targets if t.id == target_id), None)
    if target is None:
        return HTMLResponse(
            '<div class="alert alert-danger">Target no encontrado</div>',
            status_code=404,
        )

    config.max_results_per_target = max_results

    try:
        result = await scrape_target(config, target, query)
    except Exception as e:
        logger.exception("Error scraping %s", target_id)
        return templates.TemplateResponse(
            "partials/tab_results.html",
            {
                "request": request,
                "target_name": target.name,
                "records": [],
                "error": f"Error al consultar: {str(e)}",
            },
        )

    return templates.TemplateResponse(
        "partials/tab_results.html",
        {
            "request": request,
            "target_name": result.target_name,
            "records": result.records,
            "error": result.error,
        },
    )


@app.get("/api/export")
async def export_excel(
    query: str = Query(..., min_length=1),
    max_results: int = Query(50, ge=10, le=200),
):
    config.max_results_per_target = max_results

    results = await scrape_all_targets(config, query)

    excel_bytes = generate_excel(query, results)

    filename = f"metavufind_{query.replace(' ', '_')}.xlsx"

    return StreamingResponse(
        excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
