import argparse
import csv
import re

from pathlib import Path
from openpyxl import load_workbook

def csvThis(filename, sheet, delimiter, output=None):
    """
    Use the openpyxl library to open and read our xlsx
    clean it of some trash characters or spaces
    and convert it to CSV
    """
    wb = load_workbook(filename.as_posix())
    ws = wb[sheet]
    if not output:
        output = "test"
    csv_file = open("{filename}.csv".format(filename=output), 'w')
    writer = csv.writer(csv_file, delimiter=delimiter, quoting=csv.QUOTE_ALL)

    for row in tuple(ws.rows):
        row_to_write = []
        for cell in row:
            value = cell.value
            if value and type(value) != int:
                if type(value) == bytes:
                    value = value.decode("utf-8")
                value = re.sub(r"[\r\n\t\x07\x0b]", "", value)
            row_to_write.append(value)
        writer.writerow(row_to_write)

    csv_file.close()

if __name__ == "__main__":
    """
    Parse some parameters
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--xlsx_file", required=True, help="specify the file to convert")
    parser.add_argument("-s", "--sheet", required=True, help="specify the sheet to convert")
    parser.add_argument("-d", "--delimiter", default=",", help="specify a delimiter for the csv")
    parser.add_argument("-o", "--output", help="specify output file name, if not the same name as the source will be used")
    args = parser.parse_args()
    xlsx_file = Path(args.xlsx_file)
    if xlsx_file.exists() and xlsx_file.suffix == '.xlsx':
        csvThis(filename=xlsx_file, sheet=args.sheet, delimiter=args.delimiter, output=args.output)
    else:
        raise ValueError("File does not exists or is not a xlsx")
