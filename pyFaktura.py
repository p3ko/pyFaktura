import argparse
import locale
import sys
from datetime import datetime, timedelta
from decimal import Decimal
from pathlib import Path

import yaml
from borb.pdf import Document
from borb.pdf import PDF
from borb.pdf import Page
from borb.pdf.canvas.color.color import X11Color
from borb.pdf.canvas.font.font import Font
from borb.pdf.canvas.font.simple_font.true_type_font import TrueTypeFont
from borb.pdf.canvas.layout.layout_element import Alignment
from borb.pdf.canvas.layout.page_layout.multi_column_layout import SingleColumnLayout
from borb.pdf.canvas.layout.table.fixed_column_width_table import FixedColumnWidthTable as Table
from borb.pdf.canvas.layout.table.flexible_column_width_table import FlexibleColumnWidthTable as FlexibleTable
from borb.pdf.canvas.layout.table.table import TableCell
from borb.pdf.canvas.layout.text.paragraph import Paragraph
from num2words import num2words


def _build_creation_date_table() -> Table:
    table_001 = Table(number_of_rows=2, number_of_columns=2)
    font_size = Decimal(9)

    table_001.add(
        Paragraph(
            f"Data wystawienia {create_date}, {config['seller']['city']}",
            font=font,
            font_size=font_size,
            horizontal_alignment=Alignment.LEFT
        )
    )

    table_001.add(Paragraph(" "))

    table_001.add(
        Paragraph(
            f"Data sprzedaży/wykonania usługi {services_delivery_date}",
            font=font,
            font_size=font_size,
            horizontal_alignment=Alignment.LEFT
        )
    )

    table_001.add(Paragraph(" "))
    table_001.no_borders()

    return table_001


def _build_seller_buyer_table() -> Table:
    table_001 = Table(number_of_rows=5, number_of_columns=3)
    font_header_size = Decimal(10)
    font_size = Decimal(9)

    table_001.add(Paragraph("SPRZEDAWCA:", font=bold_font, font_size=font_header_size))
    table_001.add(Paragraph(" "))
    table_001.add(Paragraph("NABYWCA:", font=bold_font, font_size=font_header_size))

    table_001.add(Paragraph(config["seller"]["name"], font=font, font_size=font_size))
    table_001.add(Paragraph(" "))
    table_001.add(Paragraph(config["buyer"]["name"], font=font, font_size=font_size))

    table_001.add(Paragraph(config["seller"]["address"], font=font, font_size=font_size))
    table_001.add(Paragraph(" "))
    table_001.add(Paragraph(config["buyer"]["address"], font=font, font_size=font_size))

    table_001.add(
        Paragraph(f'{config["seller"]["zipcode"]} {config["seller"]["city"]}', font=font, font_size=font_size))
    table_001.add(Paragraph(" "))
    table_001.add(Paragraph(f'{config["buyer"]["zipcode"]} {config["buyer"]["city"]}', font=font, font_size=font_size))

    table_001.add(Paragraph(f'NIP: {config["seller"]["nip"]}', font=font, font_size=font_size))
    table_001.add(Paragraph(" "))
    table_001.add(Paragraph(f'NIP: {config["buyer"]["nip"]}', font=font, font_size=font_size))

    table_001.no_borders()

    return table_001


def _build_bank_account_table() -> Table:
    table_001 = Table(number_of_rows=1, number_of_columns=1)
    font_size = Decimal(9)
    bank_account_number = config["seller"]["bank_account_number"]

    table_001.add(Paragraph(f"NUMER RACHUNKU: {bank_account_number}", font=font, font_size=font_size))
    table_001.no_borders()

    return table_001


def _build_main_invoice_table() -> FlexibleTable:
    locale.setlocale(locale.LC_ALL, 'pl_PL.UTF-8')

    table_001 = FlexibleTable(number_of_rows=rows + 2, number_of_columns=9)
    font_size = Decimal(7.1)

    headers = [
        "L.p.",
        "Nazwa towaru/Usługi",
        "J.m.",
        "Cena Jedn. Netto",
        "Ilość",
        "Stawka VAT [%]",
        "Netto [PLN]",
        "VAT [PLN]",
        "Brutto [PLN]"
    ]

    for header in headers:
        table_001.add(
            TableCell(
                Paragraph(
                    header,
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.CENTERED
                ),
                background_color=X11Color("Gainsboro"),
                border_width=Decimal(0.5)
            )
        )

    total_netto = 0
    total_vat = 0
    total_brutto = 0
    for idx, item in enumerate(config["items"]):
        total_netto += round(item["item_netto_cost"] * item["item_quantity"], 2)
        total_vat += round((item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]), 2)
        total_brutto += round(((item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]) + (
                item["item_netto_cost"] * item["item_quantity"])), 2)
        table_001.add(
            TableCell(
                Paragraph(
                    f"{idx + 1}.",
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.CENTERED
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    item["item_name"],
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.CENTERED
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    "szt.",
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string("%.2f", item["item_netto_cost"], grouping=True).replace(u"\u202f", " "),
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string("%.2f", item["item_quantity"], grouping=True).replace(u"\u202f", " "),
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    f"{item['vat_rate']}",
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string("%.2f", item["item_netto_cost"] * item["item_quantity"],
                                         grouping=True).replace(u"\u202f", " "),
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string(
                        "%.2f",
                        (item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]),
                        grouping=True
                    ).replace(u"\u202f", " "),
                    font=font, font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string(
                        "%.2f",
                        (item["vat_rate"] / 100) *
                        (item["item_netto_cost"] * item["item_quantity"])
                        + (item["item_netto_cost"] * item["item_quantity"]),
                        grouping=True
                    ).replace(u"\u202f", " "),
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )

    for _ in range(5):
        table_001.add(
            TableCell(
                Paragraph(
                    " ",
                    font=bold_font,
                    font_size=font_size
                ),
                border_top=False,
                border_right=False,
                border_left=False,
                border_bottom=False,
                border_width=Decimal(0.5)
            )
        )
    table_001.add(
        TableCell(
            Paragraph(
                "Razem",
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )
    table_001.add(
        TableCell(
            Paragraph(
                locale.format_string("%.2f", total_netto, grouping=True).replace(u"\u202f", " "),
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )
    table_001.add(
        TableCell(
            Paragraph(
                locale.format_string("%.2f", total_vat, grouping=True).replace(u"\u202f", " "),
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )
    table_001.add(
        TableCell(
            Paragraph(
                locale.format_string("%.2f", total_brutto, grouping=True).replace(u"\u202f", " "),
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )

    table_001.set_padding_on_all_cells(Decimal(4), Decimal(4), Decimal(2), Decimal(4))

    return table_001


def _build_detailed_table() -> FlexibleTable:
    locale.setlocale(locale.LC_ALL, 'pl_PL.UTF-8')

    table_001 = FlexibleTable(number_of_rows=rows + 4, number_of_columns=4)
    font_size = Decimal(7.1)

    for _ in range(4):
        table_001.add(
            TableCell(
                Paragraph(
                    " ",
                    font=bold_font,
                    font_size=font_size
                ),
                border_top=False,
                border_right=False,
                border_left=False,
                border_bottom=False,
                border_width=Decimal(0.5)
            )
        )

    headers = [
        "Stawka VAT [%]",
        "Netto [PLN]",
        "VAT [PLN]",
        "Brutto [PLN]"
    ]

    for header in headers:
        table_001.add(
            TableCell(
                Paragraph(
                    header,
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.CENTERED
                ),
                background_color=X11Color("Gainsboro"),
                border_width=Decimal(0.5)
            )
        )

    total_netto = 0
    total_vat = 0
    total_brutto = 0
    for idx, item in enumerate(config["items"]):
        total_netto += round(item["item_netto_cost"] * item["item_quantity"], 2)
        total_vat += round((item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]), 2)
        total_brutto += round(((item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]) + (
                item["item_netto_cost"] * item["item_quantity"])), 2)

        table_001.add(
            TableCell(
                Paragraph(
                    f"{item['vat_rate']}",
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string("%.2f", item["item_netto_cost"] * item["item_quantity"],
                                         grouping=True).replace(u"\u202f", " "),
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string(
                        "%.2f",
                        (item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]),
                        grouping=True
                    ).replace(u"\u202f", " "),
                    font=font, font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )
        table_001.add(
            TableCell(
                Paragraph(
                    locale.format_string(
                        "%.2f",
                        (item["vat_rate"] / 100) *
                        (item["item_netto_cost"] * item["item_quantity"])
                        + (item["item_netto_cost"] * item["item_quantity"]),
                        grouping=True
                    ).replace(u"\u202f", " "),
                    font=font,
                    font_size=font_size,
                    horizontal_alignment=Alignment.RIGHT
                ),
                border_width=Decimal(0.5)
            )
        )

    table_001.add(
        TableCell(
            Paragraph(
                "Razem",
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )
    table_001.add(
        TableCell(
            Paragraph(
                locale.format_string("%.2f", total_netto, grouping=True).replace(u"\u202f", " "),
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )
    table_001.add(
        TableCell(
            Paragraph(
                locale.format_string("%.2f", total_vat, grouping=True).replace(u"\u202f", " "),
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )
    table_001.add(
        TableCell(
            Paragraph(
                locale.format_string("%.2f", total_brutto, grouping=True).replace(u"\u202f", " "),
                font=bold_font,
                font_size=font_size,
                horizontal_alignment=Alignment.RIGHT
            ),
            border_width=Decimal(0.5)
        )
    )

    for _ in range(4):
        table_001.add(
            TableCell(
                Paragraph(
                    " ",
                    font=bold_font,
                    font_size=font_size
                ),
                border_top=False,
                border_right=False,
                border_left=False,
                border_bottom=False,
                border_width=Decimal(0.5)
            )
        )

    table_001.set_padding_on_all_cells(Decimal(4), Decimal(4), Decimal(2), Decimal(4))

    return table_001


def _build_summary_table() -> Table:
    table_001 = Table(number_of_rows=5, number_of_columns=1)

    total_brutto = 0
    for idx, item in enumerate(config["items"]):
        total_brutto += round(((item["vat_rate"] / 100) * (item["item_netto_cost"] * item["item_quantity"]) + (
                item["item_netto_cost"] * item["item_quantity"])), 2)

    number_in_words = num2words(f'{total_brutto}', lang='pl', to='currency').replace(' euro', ' zł')

    if "centy" in number_in_words:
        number_in_words = number_in_words.replace(' centy', ' grosze')
    elif "centów" in number_in_words:
        number_in_words = number_in_words.replace(' centów', ' groszy')
    elif "cent" in number_in_words:
        number_in_words = number_in_words.replace(' cent', ' grosz')

    table_001.add(
        Paragraph(locale.format_string("Do zapłaty: %.2f PLN", total_brutto, grouping=True).replace(u"\u202f", " "),
                  font=bold_font, font_size=Decimal(10)))

    table_001.add(
        Paragraph(f"Słownie: {number_in_words}",
                  font=font, font_size=Decimal(9)))

    table_001.add(Paragraph(" "))

    table_001.add(Paragraph("Typ płatności: Przelew", font=font, font_size=Decimal(8)))

    payment_date = services_delivery_date + timedelta(days=config["time_to_pay_in_days"])
    table_001.add(Paragraph(f"Termin płatności: {payment_date.strftime('%Y-%m-%d')}", font=font, font_size=Decimal(8)))

    table_001.no_borders()

    return table_001


def _build_signature() -> Table:
    table_001 = Table(number_of_rows=1, number_of_columns=3)
    font_size = Decimal(7)

    table_001.add(
        TableCell(
            Paragraph(
                "Podpis osoby upoważnionej do wystawienia",
                font=font,
                font_size=font_size,
                horizontal_alignment=Alignment.CENTERED
            ),
            border_left=False,
            border_right=False,
            border_bottom=False,
            border_width=Decimal(0.5)
        )
    )
    table_001.add(TableCell(Paragraph(" "),
                            border_left=False,
                            border_right=False,
                            border_bottom=False,
                            border_top=False
                            )
                  )
    table_001.add(
        TableCell(
            Paragraph("Podpis osoby upoważnionej do odbioru",
                      font=font,
                      font_size=font_size,
                      horizontal_alignment=Alignment.CENTERED
                      ),
            border_left=False,
            border_right=False,
            border_bottom=False,
            border_width=Decimal(0.5)
        )
    )

    table_001.set_padding_on_all_cells(Decimal(4), Decimal(1), Decimal(1), Decimal(1))

    return table_001


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-c", "--config-file", help="Configuration file in yaml format", required=True)
    args = parser.parse_args()

    try:
        with open(args.config_file, 'r') as file:
            config = yaml.safe_load(file)
    except FileNotFoundError as error:
        print("Configuration file not exists!")
        sys.exit(1)

    rows = len(config["items"])

    # Utworzenie dokumentu
    pdf = Document()

    # Utworzenie strony
    page = Page()

    # Utworzenie layout'u
    page_layout = SingleColumnLayout(page)
    page_layout.horizontal_margin = page.get_page_info().get_width() * Decimal(0.07)
    page_layout.vertical_margin = page.get_page_info().get_height() * Decimal(0.05)

    # Dodanie tytułu
    services_delivery_date = datetime.strptime(config["services_delivery_date"], "%Y-%m-%d").date()
    create_date = datetime.strptime(config["invoice_creation_date"], "%Y-%m-%d").date()

    font_path: Path = Path(__file__).parent / config["regular_font_file"]
    font: Font = TrueTypeFont.true_type_font_from_file(font_path)

    bold_font_path: Path = Path(__file__).parent / config["bold_font_file"]
    bold_font: Font = TrueTypeFont.true_type_font_from_file(bold_font_path)

    page_layout.add(
        Paragraph(
            f"FAKTURA nr FA/{create_date.strftime('%Y')}/{create_date.strftime('%m')}/{create_date.strftime('%d')}/{config['invoice_number']}",
            font=bold_font,
            font_size=Decimal(14),
            horizontal_alignment=Alignment.CENTERED
        )
    )

    page_layout.add(Paragraph(" "))

    page_layout.add(_build_creation_date_table())

    page_layout.add(Paragraph(" "))

    # Dodanie tabeli do strony
    page_layout.add(_build_seller_buyer_table())
    page_layout.add(_build_bank_account_table())

    # Dodanie pustego paragrafu dla zrobienia miejsca
    for _ in range(2):
        page_layout.add(Paragraph(" "))

    # Dodanie tabeli z produktami do strony
    page_layout.add(_build_main_invoice_table())

    # Dodanie tabeli z produktami do strony
    page_layout.add(_build_detailed_table())
    page_layout.add(_build_summary_table())

    # Dodanie pustego paragrafu dla zrobienia miejsca
    page_layout.add(Paragraph(" "))

    # Dodanie pustego paragrafu dla zrobienia miejsca
    for _ in range(2):
        page_layout.add(Paragraph(" "))

    # Dodanie tabeli z podpisami
    page_layout.add(_build_signature())

    # Dodanie strony do dokumentu
    pdf.add_page(page)

    # Zapis dokumentu do pliku
    with open(f"FAKTURA_FA_{create_date}_{config['invoice_number']}.pdf".replace("-", "_"), "wb") as pdf_file_handle:
        PDF.dumps(pdf_file_handle, pdf)
