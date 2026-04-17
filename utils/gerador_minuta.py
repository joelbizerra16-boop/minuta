from io import BytesIO
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

WINDOWS_FONT_DIR = Path("C:/Windows/Fonts")


def _register_pdf_fonts() -> tuple[str, str]:
    preferred_fonts = [
        ("Calibri", WINDOWS_FONT_DIR / "calibri.ttf", WINDOWS_FONT_DIR / "calibrib.ttf"),
        ("Segoe UI", WINDOWS_FONT_DIR / "segoeui.ttf", WINDOWS_FONT_DIR / "segoeuib.ttf"),
        ("Arial", WINDOWS_FONT_DIR / "arial.ttf", WINDOWS_FONT_DIR / "arialbd.ttf"),
    ]

    for font_name, regular_font, bold_font in preferred_fonts:
        if regular_font.is_file() and bold_font.is_file():
            if font_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(font_name, str(regular_font)))
            bold_font_name = f"{font_name}-Bold"
            if bold_font_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(bold_font_name, str(bold_font)))
            return font_name, bold_font_name

    return "Helvetica", "Helvetica-Bold"


def _format_currency_br(value: float) -> str:
    formatted = f"{value:,.2f}"
    return f"R$ {formatted.replace(',', '_').replace('.', ',').replace('_', '.')}"


def _format_weight_br(value: float) -> str:
    formatted = f"{value:,.3f}".rstrip("0").rstrip(".")
    return formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def _format_volume_br(value: float) -> str:
    if float(value).is_integer():
        return f"{int(value):,}".replace(",", ".")
    return _format_weight_br(value)


def generate_minuta_entrega_pdf(
    records: list[dict[str, object]],
    totals: dict[str, float | int],
    numero_documento: str,
    data_emissao: str,
    transportadora: str,
    veiculo: str,
    motorista: str = "",
    placa: str = "",
    empresa: str = "BRIDA LUBRIFICANTES LTDA",
    document_title: str = "MINUTA DE ENTREGA",
    subject_label: str = "Carregamento",
) -> bytes:
    regular_font, bold_font = _register_pdf_fonts()

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=landscape(A4))

    page_width, page_height = landscape(A4)
    left_margin = 34
    right_margin = page_width - 34
    top_margin = page_height - 28
    bottom_margin = 28
    line_height = 11
    row_padding = 10
    section_gap = 18
    zebra_fill = colors.HexColor("#FAFAFA")
    light_fill = colors.HexColor("#F3F4F6")
    light_line = colors.HexColor("#DCDCDC")
    text_muted = colors.HexColor("#5B6573")

    placa_value = placa or veiculo or "--"
    motorista_value = motorista or "--"

    columns = {
        "nota": {"x": left_margin, "width": 76},
        "item": {"x": left_margin + 78, "width": 52},
        "emissao": {"x": left_margin + 132, "width": 86},
        "cliente": {"x": left_margin + 220, "width": 244},
        "cidade": {"x": left_margin + 466, "width": 116},
        "uf": {"x": left_margin + 584, "width": 34},
        "peso": {"x": left_margin + 620, "width": 68},
        "valor": {"x": left_margin + 690, "width": 82},
    }

    def wrap_text(text: object, font_name: str, font_size: int, width: float) -> list[str]:
        lines = simpleSplit(str(text or "--"), font_name, font_size, width)
        return lines or ["--"]

    def draw_label_value(x_pos: float, y_pos: float, label: str, value: object, value_offset: float = 74) -> None:
        pdf.setFont(bold_font, 10)
        pdf.setFillColor(colors.black)
        pdf.drawString(x_pos, y_pos, label)
        pdf.setFont(regular_font, 10)
        pdf.drawString(x_pos + value_offset, y_pos, str(value or "--"))

    def draw_header(continuation: bool = False) -> float:
        y_pos = top_margin
        pdf.setFillColor(colors.black)
        pdf.setFont(bold_font, 15 if not continuation else 13)
        pdf.drawString(left_margin, y_pos, empresa)
        pdf.drawRightString(right_margin, y_pos, document_title)

        pdf.setFillColor(text_muted)
        pdf.setFont(regular_font, 10)
        second_line_y = y_pos - 22
        pdf.drawString(left_margin, second_line_y, f"Emissão: {data_emissao or '--'}")
        pdf.drawString(left_margin + 220, second_line_y, f"{subject_label}: {numero_documento or '--'}")
        pdf.drawRightString(right_margin, second_line_y, f"Página: {pdf.getPageNumber()}")

        line_y = second_line_y - 10
        pdf.setStrokeColor(light_line)
        pdf.setLineWidth(0.8)
        pdf.line(left_margin, line_y, right_margin, line_y)
        return line_y - 14

    def draw_info_block(y_pos: float) -> float:
        block_height = 52
        pdf.setFillColor(colors.white)
        pdf.roundRect(left_margin, y_pos - block_height, right_margin - left_margin, block_height, 8, stroke=0, fill=1)
        pdf.setStrokeColor(light_line)
        pdf.roundRect(left_margin, y_pos - block_height, right_margin - left_margin, block_height, 8, stroke=1, fill=0)

        row_one_y = y_pos - 17
        row_two_y = y_pos - 35
        col_1_x = left_margin + 14
        col_2_x = left_margin + 390

        draw_label_value(col_1_x, row_one_y, "Transportadora", transportadora, 95)
        draw_label_value(col_2_x, row_one_y, "Veículo", veiculo or "--", 54)
        draw_label_value(col_1_x, row_two_y, "Placa", placa_value, 40)
        draw_label_value(col_2_x, row_two_y, "Motorista", motorista_value, 62)
        return y_pos - block_height - section_gap

    def draw_table_header(y_pos: float) -> float:
        header_height = 24
        pdf.setFillColor(light_fill)
        pdf.roundRect(left_margin, y_pos - header_height, right_margin - left_margin, header_height, 6, stroke=0, fill=1)
        pdf.setFillColor(colors.black)
        pdf.setFont(bold_font, 10)
        pdf.drawString(columns["nota"]["x"] + 8, y_pos - 15, "NF")
        pdf.drawString(columns["item"]["x"] + 8, y_pos - 15, "Vol")
        pdf.drawString(columns["emissao"]["x"] + 8, y_pos - 15, "Emissão")
        pdf.drawString(columns["cliente"]["x"] + 8, y_pos - 15, "Cliente")
        pdf.drawString(columns["cidade"]["x"] + 8, y_pos - 15, "Cidade")
        pdf.drawString(columns["uf"]["x"] + 8, y_pos - 15, "UF")
        pdf.drawRightString(columns["peso"]["x"] + columns["peso"]["width"] - 8, y_pos - 15, "Peso")
        pdf.drawRightString(columns["valor"]["x"] + columns["valor"]["width"] - 8, y_pos - 15, "Valor")
        return y_pos - header_height - 6

    def start_new_page() -> float:
        pdf.showPage()
        y_pos = draw_header(continuation=True)
        y_pos = draw_info_block(y_pos)
        return draw_table_header(y_pos)

    def ensure_space(y_pos: float, required_height: float) -> float:
        if y_pos - required_height < bottom_margin:
            return start_new_page()
        return y_pos

    def compute_row_height(record: dict[str, object]) -> float:
        cliente_lines = wrap_text(record.get("cliente", ""), regular_font, 9, columns["cliente"]["width"] - 16)
        cidade_lines = wrap_text(record.get("cidade", ""), regular_font, 9, columns["cidade"]["width"] - 16)
        return max(len(cliente_lines), len(cidade_lines), 1) * line_height + row_padding + 6

    current_y = draw_header()
    current_y = draw_info_block(current_y)
    current_y = draw_table_header(current_y)

    for index, record in enumerate(records):
        row_height = compute_row_height(record)
        current_y = ensure_space(current_y, row_height + 10)

        cliente_lines = wrap_text(record.get("cliente", ""), regular_font, 9, columns["cliente"]["width"] - 16)
        cidade_lines = wrap_text(record.get("cidade", ""), regular_font, 9, columns["cidade"]["width"] - 16)
        row_top = current_y - 4

        if index % 2 == 1:
            pdf.setFillColor(zebra_fill)
            pdf.roundRect(left_margin, current_y - row_height, right_margin - left_margin, row_height, 4, stroke=0, fill=1)

        pdf.setStrokeColor(light_line)
        pdf.setLineWidth(0.6)
        pdf.line(left_margin, current_y - row_height, right_margin, current_y - row_height)
        pdf.setFillColor(colors.black)
        pdf.setFont(regular_font, 10)
        pdf.drawString(columns["nota"]["x"] + 8, row_top - 10, str(record.get("nota", "--") or "--"))
        pdf.drawRightString(columns["item"]["x"] + columns["item"]["width"] - 8, row_top - 10, _format_volume_br(float(record.get("item", record.get("volume", 0.0)) or 0.0)))
        pdf.drawString(columns["emissao"]["x"] + 8, row_top - 10, str(record.get("data", "--") or "--"))

        for index, line in enumerate(cliente_lines):
            pdf.setFont(regular_font, 9)
            pdf.drawString(columns["cliente"]["x"] + 8, row_top - 10 - (index * line_height), line)

        for index, line in enumerate(cidade_lines):
            pdf.setFont(regular_font, 9)
            pdf.drawString(columns["cidade"]["x"] + 8, row_top - 10 - (index * line_height), line)

        pdf.setFont(regular_font, 10)
        pdf.drawString(columns["uf"]["x"] + 8, row_top - 10, str(record.get("uf", "--") or "--"))
        pdf.drawRightString(columns["peso"]["x"] + columns["peso"]["width"] - 8, row_top - 10, _format_weight_br(float(record.get("peso", 0.0) or 0.0)))
        pdf.drawRightString(columns["valor"]["x"] + columns["valor"]["width"] - 8, row_top - 10, _format_currency_br(float(record.get("valor", 0.0) or 0.0)))

        current_y -= row_height

    totals_block_height = 34
    footer_block_height = 96
    current_y = ensure_space(current_y, totals_block_height + footer_block_height + 12)

    pdf.setFillColor(light_fill)
    pdf.roundRect(left_margin, current_y - totals_block_height, right_margin - left_margin, totals_block_height, 8, stroke=0, fill=1)
    pdf.setFillColor(colors.black)
    pdf.setFont(bold_font, 11)
    totals_y = current_y - 22
    pdf.drawString(left_margin + 14, totals_y, f"Total Volumes: {_format_volume_br(float(totals.get('total_volumes', 0.0) or 0.0))}")
    pdf.drawString(left_margin + 220, totals_y, f"Total de NFs: {int(totals.get('total_nfs', 0) or 0)}")
    pdf.drawString(left_margin + 380, totals_y, f"Total Peso: {_format_weight_br(float(totals.get('total_peso', 0.0) or 0.0))}")
    pdf.drawRightString(right_margin - 14, totals_y, f"Total Valor: {_format_currency_br(float(totals.get('total_valor', 0.0) or 0.0))}")

    current_y -= totals_block_height + section_gap + 14
    pdf.setFillColor(colors.black)
    pdf.setFont(bold_font, 10)
    pdf.drawString(left_margin, current_y, "MERCADORIA ENTREGUE EM:")
    pdf.setFont(regular_font, 10)
    pdf.drawString(left_margin + 142, current_y, "______/______/________")
    pdf.setFont(bold_font, 10)
    pdf.drawString(left_margin + 280, current_y, "HORA DA ENTRADA:")
    pdf.setFont(regular_font, 10)
    pdf.drawString(left_margin + 392, current_y, "_____:_____")
    pdf.setFont(bold_font, 10)
    pdf.drawString(left_margin + 508, current_y, "HORA DA SAIDA:")
    pdf.setFont(regular_font, 10)
    pdf.drawString(left_margin + 606, current_y, "_____:_____")

    current_y -= 26
    pdf.setFont(regular_font, 8)
    pdf.setFillColor(text_muted)
    pdf.drawString(
        left_margin,
        current_y,
        "OBS: O PAGAMENTO DO VALOR DO FRETE NO PRAZO COMBINADO ESTARA AUTOMATICAMENTE VINCULADO AO RETORNO DO",
    )
    current_y -= 10
    pdf.drawString(
        left_margin,
        current_y,
        "CANHOTO DA NOTA FISCAL E CABECALHO DO BOLETO BANCARIO DEVIDAMENTE ASSINADO E CARIMBADO PELO CLIENTE.",
    )

    current_y -= 26
    line_left = left_margin + 20
    line_right = right_margin - 20
    customer_line_x = left_margin + ((right_margin - left_margin) / 2) + 18
    pdf.setFillColor(colors.black)
    pdf.setFont(bold_font, 10)
    pdf.line(line_left, current_y, customer_line_x - 24, current_y)
    pdf.line(customer_line_x + 24, current_y, line_right, current_y)
    pdf.setFont(regular_font, 8)
    pdf.drawCentredString((line_left + customer_line_x - 24) / 2, current_y - 12, "Transportador")
    pdf.drawCentredString((customer_line_x + 24 + line_right) / 2, current_y - 12, "Cliente / Carimbo")

    current_y -= 24
    pdf.setFont(regular_font, 8)
    pdf.setFillColor(text_muted)
    pdf.drawString(
        left_margin,
        current_y,
        "DECLARO ESTAR RETIRANDO AS MERCADORIAS REFERENTES AS NOTAS FISCAIS CONSTANTES NESTE DOCUMENTO EM PERFEITO ESTADO.",
    )

    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()