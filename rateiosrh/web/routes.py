"""HTTP routes for selecting a parser and downloading a converted workbook."""

from __future__ import annotations

import logging
import shutil
import tempfile
from io import BytesIO
from pathlib import Path

from flask import (
    Blueprint,
    current_app,
    render_template,
    request,
    send_file,
    send_from_directory,
)

from rateiosrh.core.exceptions import (
    DataFrameValidationError,
    LayoutMismatchError,
    UnknownParserError,
)
from rateiosrh.registry import list_parsers
from rateiosrh.services.conversion import ConversionService
from rateiosrh.services.excel_exporter import ExcelExporter


LOGGER = logging.getLogger("rateiosrh.web")
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

web = Blueprint("web", __name__, template_folder="templates")


def _form_response(error: str, status_code: int) -> tuple[str, int]:
    return render_template(
        "index.html", parsers=list_parsers(), error=error
    ), status_code


@web.get("/")
def index() -> str:
    return render_template("index.html", parsers=list_parsers(), error=None)


@web.get("/assets/logo_topo.png")
def logo():
    return send_from_directory(
        current_app.root_path, "logo_topo.png", mimetype="image/png"
    )


@web.get("/favicon.ico")
def favicon():
    return send_from_directory(
        current_app.root_path,
        "logo_flexivel.ico",
        mimetype="image/vnd.microsoft.icon",
    )


@web.post("/convert")
def convert():
    parser_id = request.form.get("converter_id", "").strip()
    upload = request.files.get("pdf")
    if not parser_id or upload is None or not upload.filename:
        return _form_response(
            "Selecione um conversor e um arquivo PDF para continuar.", 400
        )

    temporary_directory = Path(tempfile.mkdtemp(prefix="rateiosrh-"))
    input_path = temporary_directory / "entrada.pdf"
    output_path = temporary_directory / "resultado.xlsx"
    try:
        upload.save(input_path)
        if input_path.read_bytes()[:5] != b"%PDF-":
            shutil.rmtree(temporary_directory, ignore_errors=True)
            return _form_response("O arquivo enviado não é um PDF válido.", 400)
        result = ConversionService().convert(parser_id, input_path)
        ExcelExporter.export(result.dataframe, output_path, result.observations)
        workbook_data = output_path.read_bytes()
        shutil.rmtree(temporary_directory, ignore_errors=True)
        response = send_file(
            BytesIO(workbook_data),
            mimetype=XLSX_MIMETYPE,
            as_attachment=True,
            download_name=f"rateamento-{result.parser_id}.xlsx",
        )
        return response
    except UnknownParserError:
        shutil.rmtree(temporary_directory, ignore_errors=True)
        return _form_response("O conversor selecionado não existe.", 400)
    except LayoutMismatchError:
        shutil.rmtree(temporary_directory, ignore_errors=True)
        return _form_response(
            "O PDF é incompatível com o conversor selecionado.", 422
        )
    except DataFrameValidationError:
        shutil.rmtree(temporary_directory, ignore_errors=True)
        return _form_response(
            "O relatório não passou na validação dos dados extraídos.", 422
        )
    except Exception:
        LOGGER.exception("Falha interna durante a conversão web.")
        shutil.rmtree(temporary_directory, ignore_errors=True)
        return _form_response(
            "Ocorreu uma falha interna ao gerar o arquivo.", 500
        )
