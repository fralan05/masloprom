#!/usr/bin/env python3
"""MVP web app: upload requisites, extract fields, edit manually, generate DOCX."""
from __future__ import annotations

import cgi
import html
import io
import json
import re
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs

HOST = "0.0.0.0"
PORT = 8080

FIELD_ORDER = [
    "client.type",
    "client.name_full",
    "client.inn",
    "client.kpp",
    "client.ogrn_or_ogrnip",
    "client.legal_address",
    "client.fact_address",
    "client.bank_name",
    "client.bik",
    "client.account",
    "client.corr_account",
    "payment.form",
    "deal.amount",
    "deal.currency",
    "deal.subject",
    "contract.type",
]

CONTRACT_TEMPLATES = {
    "supply": "Договор поставки: Поставщик обязуется поставить товар, Покупатель обязуется принять и оплатить.",
    "service": "Договор оказания услуг: Исполнитель обязуется оказать услуги, Заказчик обязуется принять и оплатить.",
    "sale": "Договор купли-продажи: Продавец передает товар, Покупатель принимает и оплачивает.",
}


def parse_requisites(text: str) -> dict[str, str]:
    data: dict[str, str] = {
        "client.type": "ip" if "ИП" in text else "ooo" if "ООО" in text else "fl",
        "deal.currency": "RUB",
        "payment.form": "cash",
        "contract.type": "supply",
    }

    def find(pattern: str) -> str:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        return m.group(1).strip() if m else ""

    # Common fields
    data["client.name_full"] = find(r"(?:^|\n)\s*((?:ИП|ООО)\s[^\n]+)")
    data["client.inn"] = find(r"ИНН\s*[:№]?\s*(\d{10,12})")
    data["client.ogrn_or_ogrnip"] = find(r"ОГРНИ?П?\s*[:№]?\s*(\d{13,15})")
    data["client.account"] = find(r"(?:Номер\s*сч[её]та|р/с|расч[её]тн\w*\s*сч[её]т)\s*[:№]?\s*(\d{20})")
    data["client.corr_account"] = find(r"(?:Корсч[её]т|к/с|корр\w*\s*сч[её]т)\s*[:№]?\s*(\d{20})")
    data["client.bik"] = find(r"БИК\s*(?:банка)?\s*[:№]?\s*(\d{9})")
    data["client.bank_name"] = find(r"(?:Наименование\s*банка|Банк)\s*[:№]?\s*([^\n]+)")

    addresses = re.findall(r"\b\d{6}\b[^\n]+", text)
    data["client.legal_address"] = addresses[0].strip() if addresses else ""
    data["client.fact_address"] = addresses[1].strip() if len(addresses) > 1 else data.get("client.legal_address", "")

    return data


def validate(data: dict[str, str]) -> list[str]:
    errors = []
    required = [
        "client.name_full",
        "client.inn",
        "client.ogrn_or_ogrnip",
        "client.bank_name",
        "client.bik",
        "client.account",
        "client.corr_account",
        "payment.form",
        "deal.amount",
        "deal.currency",
    ]
    for f in required:
        if not data.get(f):
            errors.append(f"Не заполнено обязательное поле: {f}")

    inn = data.get("client.inn", "")
    if inn and not re.fullmatch(r"\d{10}|\d{12}", inn):
        errors.append("ИНН должен содержать 10 или 12 цифр")
    bik = data.get("client.bik", "")
    if bik and not re.fullmatch(r"\d{9}", bik):
        errors.append("БИК должен содержать 9 цифр")
    for f in ("client.account", "client.corr_account"):
        v = data.get(f, "")
        if v and not re.fullmatch(r"\d{20}", v):
            errors.append(f"{f} должен содержать 20 цифр")

    if data.get("client.type") == "ooo" and not data.get("client.kpp"):
        errors.append("Для ООО требуется client.kpp")

    return errors


def build_docx(data: dict[str, str]) -> bytes:
    title = "Договор (MVP)"
    template_text = CONTRACT_TEMPLATES.get(data.get("contract.type", "supply"), CONTRACT_TEMPLATES["supply"])
    lines = [
        title,
        f"Тип договора: {data.get('contract.type', 'supply')}",
        f"Форма оплаты: {data.get('payment.form', '')}",
        f"Сумма: {data.get('deal.amount', '')} {data.get('deal.currency', '')}",
        f"Предмет: {data.get('deal.subject', '')}",
        "",
        template_text,
        "",
        "Реквизиты покупателя:",
        f"{data.get('client.name_full', '')}",
        f"ИНН: {data.get('client.inn', '')}",
        f"КПП: {data.get('client.kpp', '')}",
        f"ОГРН/ОГРНИП: {data.get('client.ogrn_or_ogrnip', '')}",
        f"Юридический адрес: {data.get('client.legal_address', '')}",
        f"Фактический адрес (отгрузки): {data.get('client.fact_address', '')}",
        f"Банк: {data.get('client.bank_name', '')}",
        f"БИК: {data.get('client.bik', '')}",
        f"Р/с: {data.get('client.account', '')}",
        f"К/с: {data.get('client.corr_account', '')}",
    ]

    def p(text: str) -> str:
        esc = html.escape(text)
        return f'<w:p><w:r><w:t xml:space="preserve">{esc}</w:t></w:r></w:p>'

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" mc:Ignorable="w14">'
        '<w:body>' + ''.join(p(line) for line in lines) +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
        '</w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)
    return buf.getvalue()


def render_form(data: dict[str, str] | None = None, errors: list[str] | None = None) -> str:
    data = data or {}
    errors = errors or []

    def input_row(field: str, label: str) -> str:
        value = html.escape(data.get(field, ""))
        return f"<label>{label}<br><input name='{field}' value='{value}' style='width:100%'></label><br><br>"

    err_html = "".join(f"<li>{html.escape(e)}</li>" for e in errors)
    return f"""<!doctype html><html><head><meta charset='utf-8'><title>MVP Договоры</title></head>
<body style='font-family:Arial;max-width:860px;margin:20px auto'>
<h2>MVP: Загрузка реквизитов → извлечение → ручная правка → DOCX</h2>
<p>Поддержка форматов: txt/photo/pdf/doc/docx (для сложных форматов — частичное извлечение и ручная правка).</p>

<h3>1) Загрузка файла</h3>
<form method='post' action='/extract' enctype='multipart/form-data'>
<input type='file' name='file' required>
<button type='submit'>Загрузить и извлечь</button>
</form>

<h3>2) Ручная правка + валидация</h3>
{'<ul style="color:#b00020">'+err_html+'</ul>' if errors else ''}
<form method='post' action='/generate'>
{input_row('contract.type','Тип договора (supply/service/sale)')}
{input_row('payment.form','Форма оплаты (cash/bank_transfer/installments)')}
{input_row('deal.amount','Сумма')}
{input_row('deal.currency','Валюта')}
{input_row('deal.subject','Предмет договора')}
<hr>
{input_row('client.type','Тип клиента (ooo/ip/fl)')}
{input_row('client.name_full','Наименование клиента')}
{input_row('client.inn','ИНН')}
{input_row('client.kpp','КПП (для ООО)')}
{input_row('client.ogrn_or_ogrnip','ОГРН/ОГРНИП')}
{input_row('client.legal_address','Юридический адрес')}
{input_row('client.fact_address','Фактический адрес (отгрузки)')}
{input_row('client.bank_name','Банк')}
{input_row('client.bik','БИК')}
{input_row('client.account','Расчетный счет')}
{input_row('client.corr_account','Корреспондентский счет')}
<button type='submit'>Сгенерировать DOCX</button>
</form>
</body></html>"""


class Handler(BaseHTTPRequestHandler):
    def _send_html(self, body: str, code: int = 200) -> None:
        encoded = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def do_GET(self) -> None:  # noqa: N802
        if self.path == "/":
            self._send_html(render_form({"deal.currency": "RUB", "contract.type": "supply"}))
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        if self.path == "/extract":
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
            )
            file_item = form["file"] if "file" in form else None
            if file_item is None or getattr(file_item, "file", None) is None:
                self._send_html(render_form(errors=["Файл не загружен"]), 400)
                return
            raw = file_item.file.read()
            text = raw.decode("utf-8", errors="ignore")
            if len(text.strip()) < 20:
                text = raw.decode("cp1251", errors="ignore")
            data = parse_requisites(text)
            self._send_html(render_form(data))
            return

        if self.path == "/generate":
            length = int(self.headers.get("Content-Length", "0"))
            payload = self.rfile.read(length).decode("utf-8", errors="ignore")
            params = {k: v[0] for k, v in parse_qs(payload).items()}
            errors = validate(params)
            if errors:
                self._send_html(render_form(params, errors), 400)
                return
            docx = build_docx(params)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", "attachment; filename=generated_contract.docx")
            self.send_header("Content-Length", str(len(docx)))
            self.end_headers()
            self.wfile.write(docx)
            return

        self.send_error(HTTPStatus.NOT_FOUND)


def main() -> None:
    server = HTTPServer((HOST, PORT), Handler)
    print(f"MVP app started on http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
