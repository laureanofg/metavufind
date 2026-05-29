# MetaVufindScraping

Metabuscador academico que consulta multiples fuentes del sistema VUFIND de la Corporacion Universitaria Americana.

## Tecnologia
- Python 3.12 + FastAPI
- htmx + Bootstrap 5
- BeautifulSoup4 + httpx para scraping
- openpyxl para exportacion Excel

## Desarrollo local
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload --port 8000
```

## Docker
```bash
docker build -t metavufind .
docker run -p 8000:8000 metavufind
```
