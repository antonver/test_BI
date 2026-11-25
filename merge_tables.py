from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

HEADER = [
    "Дата и время в алк",
    "Дата и время во внешней системе",
    "Тип операции",
    "Валюта",
    "Сумма исходящего платежа",
    "Комиссия по исходящему платежу",
    "Детализация комиссии",
]
OUTPUT_FILE = "merged.xlsx"


def is_date(value: object) -> bool:
    try:
        datetime.strptime(str(value), "%d.%m.%Y %H:%M:%S")
        return True
    except ValueError:
        return False


def find_header(rows: list[tuple]) -> tuple[int | None, int | None]:
    for row_index, row in enumerate(rows):
        if HEADER[0] in row:
            return row_index, row.index(HEADER[0])
    return None, None


def read_table(path: Path) -> list[tuple]:
    sheet = load_workbook(path).active
    rows = list(sheet.iter_rows(values_only=True))

    header_row, first_column = find_header(rows)
    if header_row is None:
        return []

    table = []
    for row in rows[header_row + 1:]:
        values = row[first_column:first_column + len(HEADER)]
        if not is_date(values[0]):
            break
        table.append(values)
    return table


def main() -> None:
    result = Workbook()
    sheet = result.active
    sheet.append(HEADER)

    for path in sorted(Path(".").glob("*.xlsx")):
        if path.name == OUTPUT_FILE or path.name.startswith("~$"):
            continue
        table = read_table(path)
        print(f"{path.name}: {len(table)} rows")
        for row in table:
            sheet.append(row)

    result.save(OUTPUT_FILE)
    print(f"Saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
