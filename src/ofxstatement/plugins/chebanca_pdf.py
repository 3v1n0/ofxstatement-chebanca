from hashlib import md5
from typing import Optional, Any, List, Iterable

import logging
import os
import pathlib
import subprocess
import re

from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import StatementLine, Statement, generate_transaction_id

MEMO_TO_TYPE = {
    "Bonifico a vostro favore per ordine e conto": ("XFER", True),  # or DEP?
    "Disposizione di pagamento": ("XFER", True),
    "Storno disposizione di pagamento": ("XFER", True),
    "Giroconto": ("XFER", False),
    "Prelievo Bancomat altri Istituti": ("ATM", True),
    "Prelievo Bancomat": ("ATM", True),
    "Pagamento per utilizzo carta di credito": ("PAYMENT", True),
    "Pagamento tramite POS": ("POS", True),
    "Bonifico dall'estero": ("XFER", True),
    "Bonifico": ("XFER", True),
    "Addebito SDD": ("DIRECTDEBIT", True),
    "Pagamento imposte Delega Unificata": ("PAYMENT", True),
    "Pagamento imposte e tasse": ("FEE", False),
    "Addebito Canone": ("FEE", True),
    "Pagamenti diversi": ("PAYMENT", True),
    "Accrediti diversi": ("CREDIT", True),
    "Addebito/Accredito competenze": ("INT", False),
}


logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger("CheBancaPDF")


class CheBancaPdfParser(StatementParser):

    date_format = "%d/%m/%Y"

    def __init__(self, path_name: str) -> None:
        super().__init__()
        self.path_name = path_name

    def strip_spaces(self, string: str):
        return " ".join(string.strip().split())

    def parse_pdf_lines(self, filepath: str):
        logging.debug(f"Parsing {filepath}")

        pdftotext = subprocess.run(
            [
                "pdftotext",
                "-layout",
                filepath,
                "/dev/stdout" if os.name != "nt" else "CON",
            ],
            capture_output=True,
        )
        lines = pdftotext.stdout.decode("utf-8").split("\n")

        parsed = []
        found_first_line = False
        found_start = False

        date_regex = re.compile("^\s*\d\d/\d\d/\d\d\d\d\s.*")
        line_data = {}

        for line in lines:
            # print(line)

            if line == " ESTRATTO CONTO" or line.startswith(" CONTO DEPOSITO IN"):
                found_start = True
                continue

            if found_start and ("SALDO INIZIALE" in line or line == "OPERAZIONE"):
                found_first_line = True
                continue

            if not found_first_line:
                continue

            if "SALDO FINALE IN EURO" in line or " TOTALI" in line:
                parsed.append(line_data)
                break

            if (
                " Pagina " in line
                or " ESTRATTO CONTO dal " in line
                or "IMPOSTA DI BOLLO ASSOLTA IN MODO VIRTUALE" in line
            ):
                found_first_line = False
                continue

            if date_regex.match(line):
                splits = line.split(" ")

                if line_data:
                    parsed.append(line_data)
                    logging.debug(line_data)
                    line_data = {}

                if len(splits) < 5:
                    continue

                def next_split(i=-1):
                    for s in splits[i + 1 :]:
                        i += 1
                        if s:
                            return [splits[i], i]
                    return None

                [op_date, idx] = next_split()
                [value_date, value_date_idx] = next_split(idx)
                [amount, amount_idx] = next_split(value_date_idx)

                [_, description_idx] = next_split(amount_idx)

                line_data["op-date"] = op_date
                line_data["value-date"] = value_date
                line_data["memo"] = self.strip_spaces(
                    " ".join(splits[description_idx:])
                )

                amount_end = amount_idx + len(amount)
                value_date_end = value_date_idx + len(value_date)
                if amount_idx - value_date_end < description_idx - amount_end:
                    line_data["negative-amount"] = amount
                else:
                    line_data["amount"] = amount
            else:
                stripped = self.strip_spaces(line)
                if stripped:
                    line_data["memo"] += f" {stripped}"

        if line_data:
            logging.debug(line_data)

        return parsed

    def split_records(self) -> Iterable[List[dict]]:
        if os.path.isdir(self.path_name):
            parsed = []
            for pdf in pathlib.Path(self.path_name).glob("*.pdf"):
                parsed += self.parse_pdf_lines(pdf)
            return parsed
        else:
            return self.parse_pdf_lines(self.path_name)

    def remove_prefix(self, text, prefix):
        return text[text.lower().startswith(prefix.lower()) and len(prefix) :].lstrip()

    def parse_value(self, value: Optional[str], field: str) -> Any:
        if field == "trntype":
            for prefix, [tp, _] in MEMO_TO_TYPE.items():
                if value.lower().startswith(prefix.lower()):
                    return tp
            raise Exception("Unandled type " + value)
            # return "OTHER"

        elif field == "memo":
            for prefix, [_, strip] in MEMO_TO_TYPE.items():
                if strip:
                    value = self.remove_prefix(value, prefix)

        elif field == "negative-amount" or field == "amount":
            value = value.replace(".", "")
            if field == "negative-amount":
                field = "amount"
                value = f"-{value}"

        return super().parse_value(value, field)

    def parse_record(self, line: List[dict]) -> Optional[StatementLine]:
        amount_key = "amount" if "amount" in line else "negative-amount"
        stat_line = StatementLine(
            date=self.parse_value(line["op-date"], "date"),
            memo=self.parse_value(line["memo"], "memo"),
            amount=self.parse_value(line[amount_key], amount_key),
        )

        stat_line.date_user = self.parse_value(line["value-date"], "date_user")
        stat_line.trntype = self.parse_value(line["memo"], "trntype")

        stat_line.id = generate_transaction_id(stat_line)

        logging.debug(stat_line)
        stat_line.assert_valid()

        return stat_line

    def parse(self) -> Statement:
        reader = self.split_records()

        for line in reader:
            self.cur_record += 1
            if not line:
                continue

            parsed = self.parse_record(line)
            if parsed:
                parsed.assert_valid()

                if isinstance(parsed, StatementLine):
                    self.statement.lines.append(parsed)
                elif isinstance(parsed, InvestStatementLine):
                    self.statement.invest_lines.append(parsed)

        return self.statement


class CheBancaPdfPlugin(Plugin):
    """CheBanca Pdf"""

    def get_parser(self, filename: str) -> StatementParser:
        return CheBancaPdfParser(filename)

        # no plugin with matching signature was found
        raise Exception("No suitable CheBanca parser " "found for this statement file.")
