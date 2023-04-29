import argparse
import re
from bisect import bisect_right
from os import getcwd
from pathlib import Path
from pandas import DataFrame, to_numeric, concat
from pypdf import PdfReader

ACCOUNT_HEADER = "Girokonto"
TABLE_HEADERS = ("Buchungstag", "Vorgang", "Auftraggeber", "Buchungstext", "Ausgang")
ACCOUNT_END_SIGNAL_1 = "Neuer"
ACCOUNT_END_SIGNAL_2 = "Saldo" # both signals appear one after another = account finished
ENCODING = "cp1252"
LAST_HEADER_LEEWAY = 10 # Last column is right-aligned instead of left-aligned, so some values in this column are printed further left than the header
TITLE_FONT_SIZE_THRESHOLD = 9 # bigger than this = account title
LINE_BREAK_THRESHOLD = 11 # smaller than this = line break (not table row break)
REGEX_DATE = "\d{2}\.\d{2}\.\d{4}"

def pdfdecode(byteslist):
    return byteslist.decode(ENCODING)

def emptytable():
    return DataFrame(dtype="string")

def find_closest_header(text_x, header_xcoords):
    # E.g. if headers = [1,3,7], a text chunk with x=5 would be in the second column, or headers[1].
    # bisect_right gives us the _next_ header, which is the closest we can get in terms of convenience.
    # E.g. bisect_right([1,3,7], 5) == 2 and bisect_right([1,3,7], 3) == 2
    # If the table was right-aligned instead, bisect_left would be the correct function to use.
    index_of_next_larger_x = bisect_right(header_xcoords, text_x)
    if index_of_next_larger_x == 0: # text_x is further left than the first column header
        return None
    column_index = index_of_next_larger_x - 1
    return header_xcoords[column_index]

def parse_finanzreport(fp):
    out_of_account_parts = []
    unassignable_parts = []
    active_account = ""
    chunks_table = emptytable()
    unique_header_xcoords = set()
    end_signal_1_seen = False
    first_header = 0
    prev_col = 0.
    cur_row_y = 10000000. # y is inverted (bottom = 0, top of the page = some number around 1000)
    cur_row = 0
    skip_row = True
    def interpret_chunk(operator, operandargs, __transformation_matrix, text_matrix):
        nonlocal active_account
        nonlocal chunks_table
        nonlocal unique_header_xcoords
        nonlocal end_signal_1_seen
        nonlocal first_header
        nonlocal prev_col
        nonlocal cur_row
        nonlocal cur_row_y
        nonlocal skip_row
        if operator != b'TJ':
            return
        text = pdfdecode(operandargs[0][0])
        font_size = text_matrix[0]
        x = text_matrix[4]
        y = text_matrix[5]
        if is_account_title(text, font_size):
            active_account = text
            return
        if not active_account:
            out_of_account_parts.append(text)
            return
        if is_table_header(text):
            header_x = x if not is_last_table_header(text) else x - LAST_HEADER_LEEWAY
            unique_header_xcoords.add(header_x)
            header_xcoords = sorted(list(unique_header_xcoords))
            if is_first_table_header(text):
                first_header = x
                cur_row_y = y
            return
        if text == ACCOUNT_END_SIGNAL_1:
            end_signal_1_seen = True
            return
        if end_signal_1_seen and text == ACCOUNT_END_SIGNAL_2:
            active_account = ""
            return
        elif end_signal_1_seen:
            end_signal_1_seen = False
        column = find_closest_header(x, header_xcoords)
        if column is None or y > cur_row_y:
            unassignable_parts.append(text)
            return
        if column == first_header:
            if (cur_row_y - y) < LINE_BREAK_THRESHOLD:
                return
            if not re.match(REGEX_DATE, text):
                skip_row = True
                return
            cur_row += 1
            cur_row_y = y
            skip_row = False
        if skip_row:
            unassignable_parts.append(text)
            return
        if column == prev_col:
            chunks_table.at[cur_row, column] = f"{chunks_table.loc[cur_row, column]} {text}"
        else:
            chunks_table.at[cur_row, column] = str(text)
        prev_col = column

    def is_account_title(text, font_size):
        return text == ACCOUNT_HEADER and font_size > TITLE_FONT_SIZE_THRESHOLD

    def is_table_header(text):
        for h in TABLE_HEADERS:
            if h in text:
                return True
            
    def is_last_table_header(text):
        return TABLE_HEADERS[-1] in text
    
    def is_first_table_header(text):
        return TABLE_HEADERS[0] in text
        
    reader = PdfReader(fp, strict=True)
    pagecount = 1
    for page in reader.pages:
        unassignable_parts.append(f"Page {pagecount}: ")
        page.extract_text(visitor_operand_before=interpret_chunk)
        unassignable_parts.append("\n")
        pagecount += 1
    print(f"Extracted {len(chunks_table.index)} cash flow items from {fp.name}.")
    return chunks_table

def prettify_and_enrich_finanzreport(table, filename):
    reordered = table.reindex(sorted(table.columns), axis=1)

    print_headers = list(TABLE_HEADERS)
    print_headers[-1] = "Betrag (Originaltext)"
    headers_map = {}
    for i in range(len(print_headers)):
        headers_map[reordered.columns[i]] = print_headers[i]
    renamed = reordered.rename(columns=headers_map)

    numerified_betrag = renamed["Betrag (Originaltext)"].str.replace(".","").str.replace(",",".")
    renamed["Betrag (Zahl)"] = to_numeric(numerified_betrag)

    ## TODO implement "smart" features
    # renamed["Category"] = (figure out based on other columns - bills, groceries, rent...)

    renamed["Dateiname"] = filename
    
    return renamed
        
def write_finanzreports(tables, outfile):
    print(f"Writing table from {len(tables)} files.")
    table = concat(tables, ignore_index=True)
    table.to_csv(outfile, sep=";", encoding=ENCODING, index=False)
               
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Dieses Skript liest Comdirect-Finanzreports im pdf-Format und speichert die Umsätze vom Girokonto in einer CSV-Datei.')
    parser.add_argument('-p',dest='path',type=str,help='Pfad, in dem die Finanzreports liegen (absoluter Pfad oder Unterordner relativ zum Skript)')
    parser.add_argument('-o',dest='out',type=str,help='Ausgabe-Dateiname (Standard "girokonto.csv")')
    args = parser.parse_args()
    
    in_dir = Path(args.path) if args.path else Path(getcwd())
    wd = in_dir if in_dir.is_absolute() else getcwd()/in_dir
    collected_tables = []
    for f in wd.glob("**/Finanzreport*.pdf"):
        girokonto_table = parse_finanzreport(f)
        collected_tables.append(prettify_and_enrich_finanzreport(girokonto_table, f.name))
    out_p = args.out if args.out else "girokonto.csv"
    write_finanzreports(collected_tables, out_p)