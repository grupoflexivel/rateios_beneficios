from io import BytesIO

from app import create_app


def test_index_lists_registered_converters():
    response = create_app().test_client().get("/")

    assert response.status_code == 200
    assert b"Converter UNIMED" in response.data
    assert b"Converter Bradesco Dental" in response.data


def test_index_contains_company_branding():
    response = create_app().test_client().get("/")

    assert "Rateamento de Benefícios".encode() in response.data
    assert b"logo_topo.png" in response.data
    assert b"favicon.ico" in response.data


def test_index_is_reduced_to_branding_and_conversion_card():
    response = create_app().test_client().get("/")

    removed_copy = [
        "Do documento bruto ao detalhe que importa.",
        "Escolha o layout do relatório, envie o PDF e receba um Excel fiel aos registros de origem",
        "Transforme o relatório PDF em uma planilha pronta para conferência.",
        "Os dados são extraídos linha a linha.",
        "Saída auditável · dados textuais",
        "Uma linha por registro do PDF",
    ]
    for text in removed_copy:
        assert text.encode() not in response.data

    assert b'class="card"' in response.data
    assert b'class="intro"' not in response.data
    assert b'class="footer"' not in response.data
    assert b"background: var(--ink)" not in response.data
    assert b"box-shadow: 6px 6px 0 var(--brass)" not in response.data


def test_branding_assets_are_served():
    client = create_app().test_client()

    logo = client.get("/assets/logo_topo.png")
    favicon = client.get("/favicon.ico")

    assert logo.status_code == 200
    assert logo.mimetype == "image/png"
    assert favicon.status_code == 200
    assert favicon.mimetype in {"image/x-icon", "image/vnd.microsoft.icon"}


def test_asset_route_does_not_expose_arbitrary_root_files():
    response = create_app().test_client().get("/assets/contexto_sessao.md")

    assert response.status_code == 404


def test_convert_returns_xlsx_for_selected_parser():
    client = create_app().test_client()
    with open("FLEXIV_1166046_1_082026.pdf", "rb") as source:
        response = client.post(
            "/convert",
            data={
                "converter_id": "bradesco_dental",
                "pdf": (BytesIO(source.read()), "entrada.pdf"),
            },
            content_type="multipart/form-data",
        )

    assert response.status_code == 200
    assert response.mimetype == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert response.headers["Content-Disposition"].startswith("attachment;")


def test_convert_rejects_pdf_for_wrong_parser():
    client = create_app().test_client()
    with open("ANALITICO_TAXA_0054041451-1.PDF", "rb") as source:
        response = client.post(
            "/convert",
            data={
                "converter_id": "bradesco_dental",
                "pdf": (BytesIO(source.read()), "entrada.pdf"),
            },
            content_type="multipart/form-data",
        )

    assert response.status_code == 422
    assert "incompatível".encode() in response.data


def test_convert_rejects_invalid_pdf_as_bad_request():
    client = create_app().test_client()
    response = client.post(
        "/convert",
        data={
            "converter_id": "unimed",
            "pdf": (BytesIO(b"not a PDF"), "entrada.pdf"),
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 400


def test_convert_removes_temporary_directory_after_successful_download(
    tmp_path, monkeypatch
):
    from rateiosrh.web import routes

    temporary_directory = tmp_path / "upload-temp"

    def make_temporary_directory(prefix):
        temporary_directory.mkdir()
        return str(temporary_directory)

    monkeypatch.setattr(
        routes.tempfile,
        "mkdtemp",
        make_temporary_directory,
    )
    client = create_app().test_client()
    with open("FLEXIV_1166046_1_082026.pdf", "rb") as source:
        response = client.post(
            "/convert",
            data={
                "converter_id": "bradesco_dental",
                "pdf": (BytesIO(source.read()), "entrada.pdf"),
            },
            content_type="multipart/form-data",
        )
    response.close()

    assert response.status_code == 200
    assert not temporary_directory.exists()


def test_convert_rejects_missing_parser_or_file():
    client = create_app().test_client()

    missing_parser = client.post(
        "/convert",
        data={"pdf": (BytesIO(b"pdf"), "entrada.pdf")},
        content_type="multipart/form-data",
    )
    missing_file = client.post(
        "/convert",
        data={"converter_id": "unimed"},
        content_type="multipart/form-data",
    )

    assert missing_parser.status_code == 400
    assert missing_file.status_code == 400


def test_convert_maps_unknown_parser_to_bad_request():
    client = create_app().test_client()
    response = client.post(
        "/convert",
        data={
            "converter_id": "nao_existe",
            "pdf": (BytesIO(b"pdf"), "entrada.pdf"),
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 400


def test_convert_maps_unexpected_failure_to_internal_server_error(monkeypatch):
    from rateiosrh.web import routes

    def fail_conversion(self, parser_id, pdf_path):
        raise RuntimeError("falha simulada")

    monkeypatch.setattr(routes.ConversionService, "convert", fail_conversion)

    client = create_app().test_client()
    response = client.post(
        "/convert",
        data={
            "converter_id": "unimed",
            "pdf": (BytesIO(b"%PDF- fake"), "entrada.pdf"),
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 500
    assert b"falha interna" in response.data
