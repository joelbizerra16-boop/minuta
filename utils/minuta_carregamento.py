from dataclasses import dataclass


@dataclass(frozen=True)
class MinutaModuleConfig:
    screen_key: str
    menu_title: str
    menu_description: str
    menu_icon_key: str
    menu_button_label: str
    header_title: str
    header_subtitle: str
    subject_label: str
    summary_label: str
    export_label: str
    panel_title: str
    panel_caption: str
    pdf_title: str
    pdf_file_prefix: str


MINUTA_CARREGAMENTO_CONFIG = MinutaModuleConfig(
    screen_key="minuta",
    menu_title="Minuta de Carregamento",
    menu_description="Processamento da carga com XML, Excel e geração da minuta operacional.",
    menu_icon_key="excel",
    menu_button_label="📦 Minuta de Carregamento",
    header_title="Minuta de Carregamento",
    header_subtitle="Processamento de carga com XMLs e rotas",
    subject_label="Carregamento",
    summary_label="Resumo da Carga",
    export_label="Exportacao",
    panel_title="Painel de Notas e Itens",
    panel_caption="Visualizacao consolidada da carga com detalhamento operacional por nota fiscal.",
    pdf_title="MINUTA DE CARREGAMENTO",
    pdf_file_prefix="minuta_carregamento",
)


MINUTA_MODULES = {
    MINUTA_CARREGAMENTO_CONFIG.screen_key: MINUTA_CARREGAMENTO_CONFIG,
}
