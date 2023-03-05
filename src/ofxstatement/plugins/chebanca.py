import logging

from decimal import Decimal
from enum import Enum
from typing import Any, Iterable, Optional

from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import (
    Currency,
    Statement,
    StatementLine,
    generate_transaction_id,
)

from openpyxl import load_workbook
from openpyxl.cell import Cell

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger("CheBanca")

TYPE_MAPPING = {
    "Accrediti diversi": "CREDIT",
    "Addebito Canone": "FEE",
    "Addebito canone": "FEE",
    "Addebito Carta": "PAYMENT",
    "Addebito SDD": "DIRECTDEBIT",
    "Addebito SDD": "DIRECTDEBIT",
    "Addebito/Accredito competenze": "INT",
    "Bancomat": "ATM",
    "Bonif. v/fav.": "XFER",
    "Bonifico a vostro favore per ordine e conto": "XFER",
    "Bonifico dall'estero": "XFER",
    "Bonifico": "XFER",
    "Carta Credito.": "PAYMENT",
    "Delega Unica": "PAYMENT",
    "Disposizione di pagamento": "XFER",
    "Disposizione": "XFER",
    "Giroconto": "XFER",
    "Pagam. POS": "POS",
    "Pagamenti diversi": "PAYMENT",
    "Pagamento imposte Delega Unificata": "PAYMENT",
    "Pagamento imposte e tasse": "FEE",
    "Pagamento per utilizzo carta di credito": "PAYMENT",
    "Pagamento tramite POS": "POS",
    "Prelievo Bancomat altri Istituti": "ATM",
    "Prelievo Bancomat": "ATM",
    "Storno disposizione di pagamento": "XFER",
}


class Fields(Enum):
    DATE = "Data contabile"
    USER_DATE = "Data valuta"
    TYPE = "Tipologia"
    IN = "Entrate"
    OUT = "Uscite"
    CURRENCY = "Divisa"


class CheBancaParser(StatementParser[str]):
    date_format = "%d/%m/%Y"

    def __init__(self, filename: str) -> None:
        super().__init__()
        self.filename = filename

        logging.debug(f"Loading {self.filename}")
        self._ws = load_workbook(self.filename).active
        self._fields_to_row = {}

    def parse(self) -> Statement:
        found = False

        fields_values = [f.value.lower() for f in Fields]
        for row in self._ws:
            for cell in row:
                if isinstance(cell.value, str) and (
                    cell.value.lower() in fields_values
                ):
                    start_row = cell.row
                    start_column = cell.col_idx - 1
                    found = True
                    break

            if found:
                break

        if not found:
            raise ValueError("No 'Data contabile' cell found")

        logging.debug(
            "Statement table start cell found at "
            f"{self._ws[start_row][start_column].coordinate}"
        )

        for field in Fields:
            for cell in self._ws[start_row][start_column:]:
                if cell.value == field.value:
                    self._fields_to_row[field] = cell.col_idx - start_column - 1
                    break

        logging.debug(f"Statement table mapping are {self._fields_to_row}")

        if not [Fields.DATE, Fields.USER_DATE] & self._fields_to_row.keys():
            raise ValueError("No date column found")

        if not [Fields.IN, Fields.OUT] & self._fields_to_row.keys():
            raise ValueError("No amount column found")

        if not Fields.TYPE in self._fields_to_row.keys():
            raise ValueError("No type column")

        self._start_row = start_row + 1
        self._start_column = start_column

        return super().parse()

    def split_records(self) -> Iterable[Iterable[Cell]]:

        cells = []
        row = self._start_row
        while True:
            line_contents = self._ws[row][self._start_column :]

            if not any(cell.value for cell in line_contents):
                break

            cells.append(line_contents)
            row += 1

        return cells

    def get_field_record(self, cells: Iterable[Cell], field: Fields) -> Any:
        if field not in self._fields_to_row.keys():
            return None

        return cells[self._fields_to_row[field]].value

    def strip_spaces(self, string: str):
        return " ".join(string.strip().split())

    def parse_value(self, value: Optional[str], field: str) -> Any:
        if field == "trntype":
            native_type = value.split(" - ", 1)[0].strip()
            trntype = TYPE_MAPPING.get(native_type)

            if not trntype:
                logger.warning(f"Mapping not found for {value}")
                return "OTHER"

            return trntype

        elif field == "memo":
            try:
                return self.strip_spaces(value.split(" - ", 1)[1])
            except:
                pass

        if field == "amount" and isinstance(value, float):
            return Decimal(value)

        return super().parse_value(value, field)

    def parse_record(self, cells: Iterable[Cell]) -> StatementLine:
        stat_line = StatementLine(
            date=self.parse_value(
                self.get_field_record(cells, Fields.DATE)
                or self.get_field_record(cells, Fields.USER_DATE),
                "date",
            ),
            memo=self.parse_value(self.get_field_record(cells, Fields.TYPE), "memo"),
            amount=self.parse_value(
                self.get_field_record(cells, Fields.IN)
                or self.get_field_record(cells, Fields.OUT),
                "amount",
            ),
        )

        stat_line.date_user = self.parse_value(
            self.get_field_record(cells, Fields.USER_DATE), "date"
        )
        stat_line.trntype = self.parse_value(
            self.get_field_record(cells, Fields.TYPE), "trntype"
        )

        currency = self.parse_value(
            self.get_field_record(cells, Fields.CURRENCY), "currency"
        )

        if currency:
            stat_line.currency = Currency(symbol=currency)

        stat_line.id = generate_transaction_id(stat_line)

        logging.debug(stat_line)
        stat_line.assert_valid()

        return stat_line


class CheBancaPlugin(Plugin):
    """CheBanca! parser"""

    def get_parser(self, filename: str) -> CheBancaParser:
        return CheBancaParser(filename)
