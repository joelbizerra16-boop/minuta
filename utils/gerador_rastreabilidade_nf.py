from __future__ import annotations

from datetime import datetime, timezone
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas

from carregamentos.models.rastreabilidade_nf import RastreabilidadeNfRelatorio
from ui.assets import get_brand_logo_path
from utils.gerador_minuta import _format_weight_br
from utils.pdf_fonts import register_pdf_fonts as _register_pdf_fonts

DOCUMENT_TITLE = "RELATORIO DE RASTREABILIDADE DA NOTA FISCAL"
HEADER_LEFT_WIDTH = 148.0
HEADER_RIGHT_WIDTH = 108.0
HEADER_LOGO_MAX_WIDTH = 72.0
HEADER_LOGO_MAX_HEIGHT = 30.0


class _NumberedCanvas(canvas.Canvas):
    """Segunda passagem para numeracao Pagina X/Y no cabecalho."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._page_states: list[dict] = []

    def showPage(self) -> None:
        self._page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        total_pages = len(self._page_states) + 1
        for page_index, page_state in enumerate(self._page_states, start=1):
            self.__dict__.update(page_state)
            self._draw_page_number_overlay(page_index, total_pages)
            canvas.Canvas.showPage(self)
        self._draw_page_number_overlay(total_pages, total_pages)
        canvas.Canvas.save(self)

    def _draw_page_number_overlay(self, current_page: int, total_pages: int) -> None:
        if not hasattr(self, "_page_number_coords"):
            return
        x_pos, y_pos = self._page_number_coords
        regular_font, _ = self._page_number_fonts
        self.setFillColor(colors.HexColor("#5B6573"))
        self.setFont(regular_font, 9)
        self.drawRightString(x_pos, y_pos, f"{current_page}/{total_pages}")


def generate_rastreabilidade_nf_pdf(relatorio: RastreabilidadeNfRelatorio) -> bytes:
    regular_font, bold_font = _register_pdf_fonts()
    buffer = BytesIO()
    pdf = _NumberedCanvas(buffer, pagesize=A4)
    pdf._page_number_fonts = (regular_font, bold_font)

    page_width, page_height = A4
    left_margin = 34
    right_margin = page_width - 34
    top_margin = page_height - 28
    bottom_margin = 36
    content_width = right_margin - left_margin
    center_x = left_margin + HEADER_LEFT_WIDTH
    center_width = content_width - HEADER_LEFT_WIDTH - HEADER_RIGHT_WIDTH
    right_col_x = right_margin
    line_height = 11
    section_gap = 16
    light_fill = colors.HexColor("#F3F4F6")
    light_line = colors.HexColor("#DCDCDC")
    text_muted = colors.HexColor("#5B6573")
    highlight_fill = colors.HexColor("#FFF7E6")

    emitido_local = relatorio.emitido_em.astimezone(timezone.utc)
    emitido_data = emitido_local.strftime("%d/%m/%Y")
    emitido_hora = emitido_local.strftime("%H:%M")
    emitido_label = f"Emissão: {emitido_data} {emitido_hora}"

    state = {"page_footer_user": relatorio.emitido_por or "--"}

    def wrap(text: object, font_name: str, font_size: int, width: float) -> list[str]:
        lines = simpleSplit(str(text or "--"), font_name, font_size, width)
        return lines or ["--"]

    def draw_footer() -> None:
        pdf.setFont(regular_font, 8)
        pdf.setFillColor(text_muted)
        footer_y = 18
        pdf.drawString(left_margin, footer_y, "Sistema BRIDA")
        pdf.drawCentredString(page_width / 2, footer_y, "Relatorio de Rastreabilidade")
        pdf.drawRightString(
            right_margin,
            footer_y,
            f"{state['page_footer_user']} | {emitido_data} {emitido_hora} | Página {pdf.getPageNumber()}",
        )

    def _draw_left_region(top_y: float) -> float:
        logo_bottom = top_y
        logo_path = get_brand_logo_path()
        if logo_path is not None and logo_path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
            try:
                image = ImageReader(str(logo_path))
                img_w, img_h = image.getSize()
                scale = min(HEADER_LOGO_MAX_WIDTH / img_w, HEADER_LOGO_MAX_HEIGHT / img_h, 1.0)
                draw_w = img_w * scale
                draw_h = img_h * scale
                pdf.drawImage(
                    image,
                    left_margin,
                    top_y - draw_h + 2,
                    width=draw_w,
                    height=draw_h,
                    preserveAspectRatio=True,
                    mask="auto",
                )
                logo_bottom = top_y - draw_h - 6
            except OSError:
                logo_bottom = top_y

        empresa_lines = wrap(relatorio.empresa, bold_font, 9, HEADER_LEFT_WIDTH - 4)
        empresa_y = logo_bottom
        pdf.setFillColor(colors.black)
        for line in empresa_lines[:2]:
            pdf.setFont(bold_font, 9)
            pdf.drawString(left_margin, empresa_y, line)
            empresa_y -= 12
        return empresa_y

    def _draw_center_region(top_y: float) -> float:
        title_size = 11
        title_lines = wrap(DOCUMENT_TITLE, bold_font, title_size, center_width - 8)
        title_y = top_y
        pdf.setFillColor(colors.black)
        for line in title_lines:
            pdf.setFont(bold_font, title_size)
            pdf.drawCentredString(center_x + (center_width / 2), title_y, line)
            title_y -= 13

        pdf.setFillColor(text_muted)
        pdf.setFont(regular_font, 9)
        emission_y = title_y - 4
        pdf.drawCentredString(center_x + (center_width / 2), emission_y, emitido_label)
        return emission_y

    def _draw_right_region(top_y: float) -> float:
        label_size = 8
        value_size = 9
        gap = 11
        block_gap = 6
        y_cursor = top_y

        def draw_labeled_block(label: str, value: str, *, store_coords: bool = False) -> None:
            nonlocal y_cursor
            pdf.setFillColor(colors.black)
            pdf.setFont(bold_font, label_size)
            pdf.drawRightString(right_col_x, y_cursor, f"{label}:")
            y_cursor -= gap
            pdf.setFillColor(text_muted if label == "Página" else colors.black)
            pdf.setFont(regular_font, value_size)
            if store_coords:
                pdf._page_number_coords = (right_col_x, y_cursor)
                pdf.drawRightString(right_col_x, y_cursor, str(pdf.getPageNumber()))
            else:
                pdf.drawRightString(right_col_x, y_cursor, value)
            y_cursor -= gap + block_gap

        draw_labeled_block("Usuário", relatorio.emitido_por or "--")
        draw_labeled_block("NF", str(relatorio.resumo.numero_nf))
        draw_labeled_block("Página", "", store_coords=True)
        return y_cursor

    def draw_header() -> float:
        top_y = top_margin
        left_bottom = _draw_left_region(top_y)
        center_bottom = _draw_center_region(top_y)
        right_bottom = _draw_right_region(top_y)

        divider_y = min(left_bottom, center_bottom, right_bottom) - 10
        pdf.setStrokeColor(light_line)
        pdf.setLineWidth(0.8)
        pdf.line(left_margin, divider_y, right_margin, divider_y)
        return divider_y - 18

    def ensure_space(y_pos: float, required: float) -> tuple[float, bool]:
        if y_pos - required < bottom_margin:
            draw_footer()
            pdf.showPage()
            return draw_header(), True
        return y_pos, False

    def draw_section_title(y_pos: float, title: str) -> float:
        y_pos, _ = ensure_space(y_pos, 28)
        pdf.setFillColor(colors.black)
        pdf.setFont(bold_font, 11)
        pdf.drawString(left_margin, y_pos, title)
        y_pos -= 8
        pdf.setStrokeColor(light_line)
        pdf.line(left_margin, y_pos, right_margin, y_pos)
        return y_pos - 10

    def draw_key_values(y_pos: float, pairs: list[tuple[str, str]], columns: int = 2) -> float:
        col_width = (right_margin - left_margin) / columns
        row_height = 14
        rows = (len(pairs) + columns - 1) // columns
        y_pos, _ = ensure_space(y_pos, rows * row_height + 4)
        for index, (label, value) in enumerate(pairs):
            col = index % columns
            row = index // columns
            x_pos = left_margin + (col * col_width)
            line_y = y_pos - (row * row_height)
            pdf.setFont(bold_font, 9)
            pdf.setFillColor(colors.black)
            pdf.drawString(x_pos, line_y, f"{label}:")
            pdf.setFont(regular_font, 9)
            pdf.drawString(x_pos + 92, line_y, str(value or "--"))
        return y_pos - (rows * row_height) - section_gap

    def draw_table(
        y_pos: float,
        headers: list[str],
        rows: list[list[str]],
        *,
        widths: list[float] | None = None,
        highlight_rows: set[int] | None = None,
    ) -> float:
        if not rows:
            y_pos, _ = ensure_space(y_pos, 24)
            pdf.setFont(regular_font, 9)
            pdf.setFillColor(text_muted)
            pdf.drawString(left_margin, y_pos, "Nenhum registro.")
            return y_pos - section_gap

        table_width = right_margin - left_margin
        if widths is None:
            col_width = table_width / len(headers)
            widths = [col_width] * len(headers)

        header_height = 22

        def draw_table_header(at_y: float) -> float:
            pdf.setFillColor(light_fill)
            pdf.rect(left_margin, at_y - header_height, table_width, header_height, stroke=0, fill=1)
            pdf.setFillColor(colors.black)
            pdf.setFont(bold_font, 8)
            x_cursor = left_margin + 4
            for header, width in zip(headers, widths):
                pdf.drawString(x_cursor, at_y - 14, header)
                x_cursor += width
            return at_y - header_height - 2

        y_pos = draw_table_header(y_pos)

        for row_index, row in enumerate(rows):
            cell_lines = [wrap(cell, regular_font, 8, width - 8) for cell, width in zip(row, widths)]
            row_height = max(len(lines) for lines in cell_lines) * line_height + 8
            y_pos, page_break = ensure_space(y_pos, row_height + 4)
            if page_break:
                y_pos = draw_table_header(y_pos)

            if highlight_rows and row_index in highlight_rows:
                pdf.setFillColor(highlight_fill)
                pdf.rect(left_margin, y_pos - row_height, table_width, row_height, stroke=0, fill=1)

            pdf.setStrokeColor(light_line)
            pdf.setLineWidth(0.5)
            pdf.line(left_margin, y_pos - row_height, right_margin, y_pos - row_height)

            pdf.setFillColor(colors.black)
            x_cursor = left_margin + 4
            for col_index, (lines, width) in enumerate(zip(cell_lines, widths)):
                for line_index, line in enumerate(lines):
                    pdf.setFont(regular_font, 8)
                    pdf.drawString(x_cursor, y_pos - 10 - (line_index * line_height), line)
                x_cursor += width
            y_pos -= row_height

        return y_pos - section_gap

    resumo = relatorio.resumo
    current_y = draw_header()
    current_y = draw_section_title(current_y, "RESUMO DA NF")
    current_y = draw_key_values(
        current_y,
        [
            ("Numero da NF", resumo.numero_nf),
            ("Chave da NF", resumo.chave_nfe),
            ("Destinatario", resumo.destinatario),
            ("Quantidade de itens", str(resumo.quantidade_itens)),
            ("Peso total", _format_weight_br(resumo.peso_total)),
            ("Quantidade de carregamentos", str(resumo.quantidade_carregamentos)),
            ("Quantidade de reentregas", str(resumo.quantidade_reentregas)),
            ("Primeira saida", resumo.primeira_saida),
            ("Ultima saida", resumo.ultima_saida),
            ("Status atual", resumo.status_atual),
        ],
    )

    current_y = draw_section_title(current_y, "HISTORICO OPERACIONAL")
    historico_headers = [
        "Data/Hora",
        "Carregamento",
        "Modalidade",
        "Status",
        "Usuário",
        "Motorista",
        "Veículo",
        "Placa",
        "Rota",
        "PDF",
        "Documento",
    ]
    historico_rows = []
    highlight_rows: set[int] = set()
    for index, linha in enumerate(relatorio.historico):
        if linha.reentrega or linha.balcao:
            highlight_rows.add(index)
        historico_rows.append(
            [
                linha.data_hora.astimezone(timezone.utc).strftime("%d/%m/%Y %H:%M"),
                linha.numero_carregamento,
                linha.modalidade,
                linha.status,
                linha.usuario,
                linha.motorista,
                linha.veiculo,
                linha.placa,
                linha.rota,
                linha.pdf_gerado,
                linha.documento,
            ]
        )
    historico_widths = [62, 58, 48, 48, 42, 48, 38, 38, 48, 34, 42]
    current_y = draw_table(current_y, historico_headers, historico_rows, widths=historico_widths, highlight_rows=highlight_rows)

    if relatorio.reentregas:
        current_y = draw_section_title(current_y, "REENTREGAS")
        current_y = draw_table(
            current_y,
            ["Data", "Carregamento", "Usuário", "Motivo", "Status"],
            [[item.data, item.carregamento, item.usuario, item.motivo, item.status] for item in relatorio.reentregas],
            widths=[70, 80, 70, 180, 70],
        )

    current_y = draw_section_title(current_y, "VEÍCULOS UTILIZADOS")
    current_y = draw_table(
        current_y,
        ["Veículo", "Placa", "Qtd. viagens", "Motorista"],
        [
            [item.veiculo, item.placa, str(item.quantidade_viagens), item.motorista]
            for item in relatorio.veiculos
        ],
        widths=[120, 80, 80, 200],
    )

    current_y = draw_section_title(current_y, "USUARIOS ENVOLVIDOS")
    current_y = draw_table(
        current_y,
        ["Usuario", "Qtd. operacoes", "Primeira operacao", "Ultima operacao"],
        [
            [item.usuario, str(item.quantidade_operacoes), item.primeira_operacao, item.ultima_operacao]
            for item in relatorio.usuarios
        ],
        widths=[90, 80, 120, 120],
    )

    current_y = draw_section_title(current_y, "MODALIDADES")
    current_y = draw_table(
        current_y,
        ["Modalidade", "Quantidade"],
        [[item.modalidade, str(item.quantidade)] for item in relatorio.modalidades],
        widths=[200, 80],
    )

    current_y = draw_section_title(current_y, "DOCUMENTOS GERADOS")
    current_y = draw_table(
        current_y,
        ["Minuta", "Romaneio", "Data", "Usuario", "Impressoes", "Ultima impressao"],
        [
            [
                item.minuta,
                item.romaneio,
                item.data,
                item.usuario,
                str(item.quantidade_impressoes),
                item.ultima_impressao,
            ]
            for item in relatorio.documentos
        ],
        widths=[50, 58, 72, 70, 58, 90],
    )

    if relatorio.estatisticas:
        stats = relatorio.estatisticas
        current_y = draw_section_title(current_y, "ESTATISTICAS")
        current_y = draw_key_values(
            current_y,
            [
                ("Total de carregamentos", str(stats.total_carregamentos)),
                ("Total de itens expedidos", str(stats.total_itens_expedidos)),
                ("Peso expedido", _format_weight_br(stats.peso_expedido)),
                ("Reentregas", str(stats.total_reentregas)),
                ("Retiradas em balcao", str(stats.total_balcao)),
                ("Veículos diferentes", str(stats.veiculos_diferentes)),
                ("Motoristas diferentes", str(stats.motoristas_diferentes)),
                ("Usuários envolvidos", str(stats.usuarios_envolvidos)),
            ],
        )

    current_y = draw_section_title(current_y, "LINHA DO TEMPO")
    timeline_lines = []
    for index, evento in enumerate(relatorio.timeline):
        if index > 0:
            timeline_lines.append(["", "↓", ""])
        timeline_lines.append(
            [
                evento.data_hora.astimezone(timezone.utc).strftime("%d/%m/%Y %H:%M"),
                evento.rotulo,
            ]
        )
    current_y = draw_table(
        current_y,
        ["Data/Hora", "Evento"],
        timeline_lines,
        widths=[90, 400],
    )

    draw_footer()
    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()
