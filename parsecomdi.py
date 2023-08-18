import argparse
import re
from bisect import bisect_right
from os import getcwd
from pathlib import Path
from statistics import median
from pandas import DataFrame, to_numeric, concat
from numpy import nan
from pypdf import PdfReader

ACCOUNT_TITLES = ("Girokonto", "Visa-Karte")
TABLE_HEADERS = ("Buchungstag", "Vorgang", "Auftraggeber", "Buchungstext", "Ausgang")
IGNORED_SUBHEADERS = ("Valuta", "Umsatztag")
PRINT_HEADERS = ("Buchungstag (orig)", "Vorgang (orig)", "Auftraggeber (orig)", "Buchungstext (orig)", "Ausgang/Eingang (orig)")
ACCOUNT_END = "Neuer Saldo" # 2014 reports and older present the phrase as one text chunk
ACCOUNT_END_SIGNAL_1 = "Neuer" # newer reports have the two words in separate chunks
ACCOUNT_END_SIGNAL_2 = "Saldo"
ENCODING = "cp1252"
MISALIGNED_TABLE_LEEWAY = 1 # 2014 and older reports don't quite align all-caps texts in the Girokonto table correctly with their column header
LAST_HEADER_LEEWAY = 10 # Last column is right-aligned instead of left-aligned, so some values in this column are printed further left than the header
# FIXME: There's a possibility that LAST_HEADER_LEEWAY needs to be slightly bigger for reports with six-figure transactions and up
TITLE_FONT_SIZE_THRESHOLD = 9 # bigger than or equal this = account title
ANY_FONT_SIZE_THRESHOLD = 3 # smaller than this = illegible text (to be ignored) or bug
LINE_BREAK_THRESHOLD = 12.01 # smaller than this = line break (not table row break)
REGEX_DATE = "A?\d{2}\.\d{2}\.\d{4}" # for some reason, in at least one instance, the date said "A29.11.2019"
REGEX_IBANBIC = "(?P<sender>.+?) (?P<ibanbic>[A-Z]{2}\d{2}[A-Z0-9]+ [A-Z]{6}[A-Z0-9]{5})" # https://stackoverflow.com/questions/21928083/iban-validation-check
REGEX_REFTEXT = "(?P<manualref>.*?) ?End-to-End-Ref\.:(?P<endref>.*)"
END_WORD = "(?:[^a-zA-Z]|$)"
CATEGORY_REGEX = [
    ("Vorgang (orig)", "(?i:wertpapiere)", "Depot", "Kauf / Verkauf"),
    ("Vorgang (orig)", "(?i:Termingeld)", "Depot", "Kauf / Verkauf"),
    ("Vorgang (orig)", "(?:Bar)", "Einkommen", "Bargeldeinzahlung"),
    ("Vorgang (orig)", "(?i:auszahlung gaa)|(?i:barauszahlung)", "Sonstige Ausgaben", "Bargeldauszahlung"),
    ("Auftraggeber-Name", "comdirect Visa", "Ausschließen", "Interner Übertrag"),
    ("Auftraggeber-Name", "(?i:stadtwerke)|(?i:abfall)", "Wohnen", "Nebenkosten"),
    ("Auftraggeber-Name", "(?i:kaufland)|(?i:lidl)|(?i:aldi)|(?i:rewe)|(?i:penny)|(?i:edeka)|(?i:tegut)", "Lebenshaltung", "Lebensmittel"),
    ("Auftraggeber-Name", "(?i:drogeriemarkt)|(?i:rossmann)", "Lebenshaltung", "Drogerie"),
    ("Auftraggeber-Name", "(?i:deichmann)|(?i:c\+a)", "Lebenshaltung", "Kleidung / Schuhe"),
    ("Auftraggeber-Name", f"(?i:obi{END_WORD})|(?i:bauhaus{END_WORD})", "Lebenshaltung", "Heim / Garten"),
    ("Auftraggeber-Name", "(?i:apotheke)", "Lebenshaltung", "Medikamente"),
    ("Auftraggeber-Name", "(?i:db vertrieb)|(?i:deutsche bahn)|(?i:rnv)", "Verkehrsmittel", "Öffentliche Verkehrsmittel"),
    ("Auftraggeber-Name", "(?:OIL)", "Verkehrsmittel", "Auto / Tanken"),
    ("Auftraggeber-Name", "(?i:unitymedia)|(?i:vodafone)|(?i:telefonica)|(?i:drillisch)|(?i:congstar)", "Digital", "Internet / Telefon"),
    ("Auftraggeber-Name", "(?i:rundfunk)", "Digital", "Rundfunksteuer"),
    ("Auftraggeber-Name", f"(?i:mcdonalds)|(?i:kfc{END_WORD})|(?i:nordsee gmbh)|(?i:moschmosch)|(?i:gastronomie)|(?i:restaurant)", "Freizeit", "Gastronomie"),
    ("Auftraggeber-Name", "(?i:cineplex)|(?i:filmpalast)", "Freizeit", "Unterhaltung / Kino / Kultur"),
    ("Auftraggeber-Name", "(?i:germanwings)", "Reisen", "Flug / Bahn / Bus / Taxi"),
    ("Auftraggeber-Name", "(?i:bundesagentur für arbeit)", "Einkommen", "Arbeitslosengeld"),
    ("Buchungsnotiz", "(?i:darlehen)", "Darlehen", "Tilgung"),
    ("Buchungsnotiz", "(?i:miete)", "Wohnen", "Miete"),
    ("Buchungsnotiz", "(?i:nebenkostenabrechnung)", "Wohnen", "Miete"),
    ("Buchungsnotiz", "(?i:hausgeld)", "Wohnen", "Hausgeld"),
    ("Buchungsnotiz", "(?i:strom)|(?i:wasser|(?i:heizung)|(?i:fernwaerme)|(?i:abfall)", "Wohnen", "Nebenkosten"),
    ("Buchungsnotiz", f"(?i:rewe{END_WORD})", "Lebenshaltung", "Lebensmittel"),
    ("Buchungsnotiz", f"(?i:zara{END_WORD})|(?i:primark)", "Lebenshaltung", "Kleidung / Schuhe"),
    ("Buchungsnotiz", "(?i:uebertrag auf)|(?i:ueberweisung auf v\s?isa-kar\s?te)|(?i:summe wochenabrechnung \s?visa)|(?i:summe monatsabrechnung visa)", "Ausschließen", "Interner Übertrag"),
    ("Buchungsnotiz", f"(?i:wage{END_WORD})|(?i:salary)|(?i:gehalt)|(?i:lohn)|(?i:bezuege )", "Einkommen", "Gehalt"),
    ("Buchungsnotiz", "(?i:ertraegnisgutschrift)", "Einkommen", "Sparen / Anlegen"),
    ("Buchungsnotiz", "(?i:bargeldeinzahlung)", "Einkommen", "Bargeldeinzahlung"),
    ("Buchungsnotiz", "(?i:bargeldauszahlung)", "Sonstige Ausgaben", "Bargeldauszahlung"),
    ("Buchungsnotiz", "(?i:apotheke)", "Lebenshaltung", "Medikamente"),
    ("Buchungsnotiz", "(?i:dbvertrieb)", "Verkehrsmittel", "Öffentliche Verkehrsmittel"), # when paying via Paypal
    ("Buchungsnotiz", "(?i:ryanair)|(?i:wizz air)|(?i:tarom)|(?i:fluegede)", "Reisen", "Flug / Bahn / Bus / Taxi"), # when paying via Paypal
    ("Buchungsnotiz", "(?i:hotel)", "Reisen", "Hotel / Unterkunft"), # when paying by VISA
    ("Buchungsnotiz", "(?i:burger king)", "Freizeit", "Gastronomie"),
    ("Buchungsnotiz", "(?i:theater)|(?i:eventim)|(?i:fkp scorpio)", "Freizeit", "Unterhaltung / Kino / Kultur"), # when paying by VISA
    ("Buchungsnotiz", "(?i:humblebundl)|(?i:steam\s?games)|(?i:wargaming)", "Freizeit", "Hobbies"), # when paying via Paypal
]

def parse_finanzreport(fp):
    out_of_account_parts = []
    unassignable_parts = []
    cur_font_size_old_pdf_format = 0 # 2014 and older reports contain font size in separate chunks from the text they apply to
    account_tables = {title:DataFrame() for title in ACCOUNT_TITLES}
    active_account = ""
    header_xcoords = {title:[] for title in ACCOUNT_TITLES}
    end_signal_1_seen = False
    first_header = 0
    prev_col = 0.
    cur_row_y = 10000000. # y is inverted (bottom = 0, top of the page = some number around 1000)
    first_col_line_breaks_y = 0.
    cur_row = 0
    row_does_not_start_with_date = True

    def interpret_chunk(operator, operandargs, __transformation_matrix, text_matrix):
        nonlocal cur_font_size_old_pdf_format
        nonlocal account_tables
        nonlocal active_account
        nonlocal header_xcoords
        nonlocal end_signal_1_seen
        nonlocal first_header
        nonlocal prev_col
        nonlocal cur_row
        nonlocal cur_row_y
        nonlocal first_col_line_breaks_y
        nonlocal row_does_not_start_with_date
        if operator != b'TJ' and operator != b'Tj':
            if operator == b'Tf':
                cur_font_size_old_pdf_format = operandargs[1]
            return
        text = pdfdecode(operandargs[0][0] if operator == b'TJ' else operandargs[0])
        font_size = text_matrix[0]
        if font_size < ANY_FONT_SIZE_THRESHOLD and cur_font_size_old_pdf_format < ANY_FONT_SIZE_THRESHOLD:
            unassignable_parts.append(text)
            return
        font_size = font_size if font_size >= ANY_FONT_SIZE_THRESHOLD else cur_font_size_old_pdf_format
        x = text_matrix[4]
        y = text_matrix[5]
        # Handle title
        if is_account_title(text, font_size):
            active_account = get_corresponding_title(text)
            first_header = 0
            prev_col = 0.
            cur_row_y = 10000000.
            cur_row = 0
            row_does_not_start_with_date = True
            return
        # Skip until title was seen
        if not active_account:
            out_of_account_parts.append(text)
            return
        table = account_tables[active_account]
        headers = header_xcoords[active_account]
        # Handle table headers
        if is_table_header(text):
            header_x = x if not is_last_table_header(text) else x - LAST_HEADER_LEEWAY
            if header_x not in headers:
                headers.append(header_x)
                headers = sorted(headers)
            if is_first_table_header(text):
                first_header = x
                cur_row_y = y
            return
        column = find_closest_header(x + LAST_HEADER_LEEWAY, headers)
        # Skip if text is outside table, or no headers seen yet
        # Also skip text chunks from higher up in the page than the current table row.
        # The report format since June 2016 throws text chunks from the page header into the middle of the table for some reason.
        # Remember that y is inverted. y=0 is the page bottom. 
        if column is None or y > cur_row_y:
            unassignable_parts.append(text)
            return
        # Handle end of Girokonto table
        if text == ACCOUNT_END_SIGNAL_1:
            end_signal_1_seen = True
            return
        if text == ACCOUNT_END or (end_signal_1_seen and text == ACCOUNT_END_SIGNAL_2):
            active_account = ""
            return
        elif end_signal_1_seen:
            end_signal_1_seen = False
        # Handle start of table row
        if column == first_header:
            # The first column contains the transaction date and the date the value was credited to the account.
            # These two dates are practically always identical, so discard the second.
            if is_regular_text_line_break(y, first_col_line_breaks_y):
                first_col_line_breaks_y = y
                # Print out in case dates are not identical (usually investment transactions, sometimes cash withdrawal)
                if text not in IGNORED_SUBHEADERS and text != table.loc[cur_row, column]:
                    print(f"Info: Wertstellung eines Vorgangs am {text} =/= Buchungstag {table.loc[cur_row, column]}.")
                return
            if not is_date(text):
                row_does_not_start_with_date = True
                return
            if text[0].isalpha():
                text = text[1:] # Fix date "A29.11.2019" found in at least one report
            cur_row += 1
            cur_row_y = first_col_line_breaks_y = y
            row_does_not_start_with_date = False
        # Skip until we have a proper table row (eliminate page header, footer, and "Alter Saldo")
        if row_does_not_start_with_date:
            unassignable_parts.append(text)
            return
        #print(int(column), int(prev_col), len(headers), text)
        if column == prev_col:
            table.at[cur_row, column] = f"{table.loc[cur_row, column]} {text}"
        else:
            table.at[cur_row, column] = text
        prev_col = column

    def pdfdecode(byteslist):
        return byteslist.decode(ENCODING)

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
    
    def is_date(text):
        return re.match(REGEX_DATE, text)

    def is_regular_text_line_break(y, cur_row_y):
        return (cur_row_y - y) < LINE_BREAK_THRESHOLD

    def is_account_title(text, font_size):
        return font_size > TITLE_FONT_SIZE_THRESHOLD and any(title in text for title in ACCOUNT_TITLES)
    
    def get_corresponding_title(text):
        return [title for title in ACCOUNT_TITLES if title in text][0]

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
        out_of_account_parts.append(f"Page {pagecount}: ")
        unassignable_parts.append(f"Page {pagecount}: ")
        # The default behaviour of extract_text misinterprets spacer signals in Finanzreport PDFs.
        # The resulting text lacks spaces and line breaks where it should have them.
        # It is impossible to determine afterwards where the resulting combined text pieces should be split.
        # So instead we use a custom interpreter that parses and collects all text chunks as they are extracted.
        page.extract_text(visitor_operand_before=interpret_chunk)
        out_of_account_parts.append("\n")
        unassignable_parts.append("\n")
        pagecount += 1
    for title, table in account_tables.items():
        if title == "Visa-Karte" and len(table.columns) == 4:
            # Old report format does not have the middle heading "Auftraggeber", insert a dummy
            table[median(header_xcoords[title])] = ""
        for x in header_xcoords[title]:
            if x not in table.columns:
                table[x] = ""
        print(f"Extracted {len(table.index)} cash flow items in {title} from {fp.name}.")
    return account_tables

def prettify_and_enrich_finanzreport(table, filename, account_title):
    reordered = table.reindex(sorted(table.columns), axis=1)
    
    headers_map = {}
    for i in range(len(PRINT_HEADERS)):
        headers_map[reordered.columns[i]] = PRINT_HEADERS[i]
    renamed = reordered.rename(columns=headers_map)

    renamed["Dateiname"] = filename
    renamed["Konto"] = account_title

    iban_named_matches: DataFrame = renamed[PRINT_HEADERS[2]].str.extract(REGEX_IBANBIC)
    renamed["Auftraggeber-Name"] = iban_named_matches["sender"].fillna(renamed[PRINT_HEADERS[2]])
    renamed["Auftraggeber-IBAN/BIC"] = iban_named_matches["ibanbic"]

    ref_named_matches: DataFrame = renamed[PRINT_HEADERS[3]].str.extract(REGEX_REFTEXT)
    renamed["Buchungsnotiz"] = ref_named_matches["manualref"].fillna(renamed[PRINT_HEADERS[3]])
    renamed["Buchungsreferenz"] = ref_named_matches["endref"]

    numerified_betrag = renamed[PRINT_HEADERS[4]].str.replace(".","").str.replace(",",".")
    renamed["Betrag (Zahl)"] = to_numeric(numerified_betrag)

    renamed["Kategorie"] = nan
    renamed["Unterkategorie"] = nan
    
    for matchcol, regex, category, subcategory in CATEGORY_REGEX:
        matches = renamed[matchcol].str.contains(regex, na=False)
        renamed["Kategorie"] = renamed["Kategorie"].where(~matches, other=category) # where() replaces False with other m(
        renamed["Unterkategorie"] = renamed["Unterkategorie"].where(~matches, other=subcategory) # where() replaces False with other m(
    
    return renamed

def write_finanzreports(tables, outfile):
    print(f"Writing final table from {len(tables)} account tables.")
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
        table_per_account: dict[DataFrame] = parse_finanzreport(f)
        if len(table_per_account) != len(ACCOUNT_TITLES):
            print(f"{abs(len(table_per_account)-len(ACCOUNT_TITLES))} account titles had no matching table. Please check why.")
            continue
        if all(table.empty for table in table_per_account.values()):
            print(f"No transations found in {f.name}. Please check if this is an error.")
            continue
        if any([not table.empty and len(table.columns) != len(TABLE_HEADERS) for table in table_per_account.values()]):
            print("At least one table does not have exactly one column per expected header. Please check why.")
            continue
        prettified_tables = [
            prettify_and_enrich_finanzreport(table, f.name, account) for account, table in table_per_account.items() if not table.empty
            ]
        [collected_tables.append(table) for table in prettified_tables]
    out_p = args.out if args.out else "girokonto.csv"
    write_finanzreports(collected_tables, out_p)