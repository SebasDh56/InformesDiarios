from pathlib import Path
import shutil
from fastapi import FastAPI, Request, UploadFile, File, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from .db import Base, engine, get_db
from .models import Cronograma, Recorrido
from .services.cronograma_service import parse_cronograma
from .services.report_service import generar_word

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

Base.metadata.create_all(bind=engine)

app = FastAPI(title="ES2027 - Generador de Informes")

app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "app" / "templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request, db: Session = Depends(get_db)):
    cronogramas = db.query(Cronograma).order_by(Cronograma.id.desc()).all()
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "cronogramas": cronogramas},
    )

@app.post("/cronograma/importar")
async def importar_cronograma(
    archivo: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    if not archivo.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Debe cargar un archivo Excel .xlsx o .xlsm")

    destino = UPLOAD_DIR / archivo.filename
    with destino.open("wb") as buffer:
        shutil.copyfileobj(archivo.file, buffer)

    data = parse_cronograma(str(destino))

    cron = Cronograma(
        filename=archivo.filename,
        delegacion=data["delegacion"],
        tipo_eleccion=data["tipo_eleccion"],
        periodo_texto=data["periodo_texto"],
    )
    db.add(cron)
    db.flush()

    for item in data["recorridos"]:
        db.add(Recorrido(cronograma_id=cron.id, **item))

    db.commit()

    return RedirectResponse(
        url=f"/informe/{cron.id}",
        status_code=303,
    )

@app.get("/informe/{cronograma_id}", response_class=HTMLResponse)
def editar_informe(
    cronograma_id: int,
    request: Request,
    db: Session = Depends(get_db),
):
    cron = db.query(Cronograma).filter(Cronograma.id == cronograma_id).first()
    if not cron:
        raise HTTPException(status_code=404, detail="Cronograma no encontrado")

    fechas = sorted(set(r.fecha for r in cron.recorridos if r.fecha))
    return templates.TemplateResponse(
        "informe.html",
        {
            "request": request,
            "cron": cron,
            "fechas": fechas,
        },
    )

@app.post("/informe/{cronograma_id}/guardar")
async def guardar_recorridos(
    cronograma_id: int,
    request: Request,
    db: Session = Depends(get_db),
):
    cron = db.query(Cronograma).filter(Cronograma.id == cronograma_id).first()
    if not cron:
        raise HTTPException(status_code=404, detail="Cronograma no encontrado")

    form = await request.form()

    for rec in cron.recorridos:
        rec.calle_inicial = (form.get(f"calle_inicial_{rec.id}") or "").strip() or None
        rec.calle_final = (form.get(f"calle_final_{rec.id}") or "").strip() or None
        rec.observaciones = (form.get(f"observaciones_{rec.id}") or "").strip() or None
        rec.hay_articulos = form.get(f"hay_articulos_{rec.id}") == "on"

    db.commit()
    return RedirectResponse(url=f"/informe/{cronograma_id}", status_code=303)

@app.post("/informe/{cronograma_id}/word")
async def generar_informe_word(
    cronograma_id: int,
    numero_informe: str = Form(...),
    db: Session = Depends(get_db),
):
    cron = db.query(Cronograma).filter(Cronograma.id == cronograma_id).first()
    if not cron:
        raise HTTPException(status_code=404, detail="Cronograma no encontrado")

    safe_num = numero_informe.replace("/", "-").replace("\\", "-").strip()
    filename = f"{safe_num or 'INFORME_ES2027'}.docx"
    output = OUTPUT_DIR / filename

    generar_word(cron, output, numero_informe=safe_num)

    return FileResponse(
        output,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
