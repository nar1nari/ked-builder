import shutil
import tempfile
import traceback
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from starlette.background import BackgroundTask
from starlette.requests import Request

from ked_builder import build_ked

app = FastAPI(title="KED Builder")

BASE_DIR = Path(__file__).parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

UPLOAD_ROOT = Path(tempfile.gettempdir()) / "ked-builder-web"
UPLOAD_ROOT.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "index.html")


def _save_upload(upload: UploadFile, dest_dir: Path) -> str:
    suffix = Path(upload.filename or "").suffix
    dest_path = dest_dir / f"{uuid.uuid4().hex}{suffix}"
    with dest_path.open("wb") as f:
        shutil.copyfileobj(upload.file, f)
    return str(dest_path)


def _cleanup(session_dir: Path) -> None:
    shutil.rmtree(session_dir, ignore_errors=True)


@app.post("/generate")
async def generate(
    year: int = Form(...),
    program_number: str = Form(...),
    program_name: str = Form(...),
    program_name_kk: str = Form(...),
    ru_doc: UploadFile = File(...),
    kz_doc: UploadFile = File(...),
    xlsx_doc: UploadFile = File(...),
    template_doc: UploadFile | None = File(None),
):
    session_dir = UPLOAD_ROOT / uuid.uuid4().hex
    session_dir.mkdir(parents=True)

    try:
        ru_path = _save_upload(ru_doc, session_dir)
        kz_path = _save_upload(kz_doc, session_dir)
        xlsx_path = _save_upload(xlsx_doc, session_dir)
        template_path = (
            _save_upload(template_doc, session_dir)
            if template_doc is not None and template_doc.filename
            else None
        )

        output_path = session_dir / "ked_output.docx"

        build_ked(
            year=year,
            program_number=program_number,
            program_name=program_name,
            program_name_kk=program_name_kk,
            xlsx_doc=xlsx_path,
            ru_doc=ru_path,
            kz_doc=kz_path,
            output_doc=str(output_path),
            template_doc=template_path,
        )

        if not output_path.exists():
            raise RuntimeError("build_ked не создал выходной файл")

        filename = f"KED_{program_number}_{year}.docx"
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
            background=BackgroundTask(_cleanup, session_dir),
        )
    except Exception as exc:
        traceback.print_exc()
        _cleanup(session_dir)
        return JSONResponse(
            status_code=400,
            content={"detail": f"Не удалось сгенерировать документ: {exc}"},
        )
