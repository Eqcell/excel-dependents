"""Very raw script to trace external references in folder with Excel files.

See Job 1-3 in __main___.

Uses Python 3.6.

"""

import csv
import itertools
from pathlib import Path 

import xlwings as xw
import pandas as pd
from openpyxl.utils import get_column_letter


ENC = 'utf-8'
CSV_FORMAT = dict(delimiter='\t', lineterminator='\n')



ROOT = Path(__file__).parent #Path("D:\\links\\базовый сценарий")
CSV_ROOT = ROOT / 'csv'

# transform cell names
def xlref(row, column, zero_indexed=False):
    if zero_indexed:
        row += 1
        column += 1
    return get_column_letter(column) + str(row)


# filters
def has_external_ref(s):
    try:
        return "\\"  in s and "]" in s
    except TypeError:
        return False

assert True == has_external_ref('d:\ДКП\[Система процентных ставок.xlsx]Ставки 51')


def get_filename(formula):
    res = formula.split("'")[1]
    if "'" in res:
        return get_filename(res)
    else:
        return res  
    
assert 'd:\DMPKA-3.1\[BDDRN цел.xls]выплаты' == get_filename(
       "='d:\DMPKA-3.1\[BDDRN цел.xls]выплаты'!B51")


# csv file access
def to_csv(rows, path):
    """Accept iterable of rows *rows* and write in to *csv_path*"""
    with path.open('w', encoding=ENC) as csvfile:
        filewriter = csv.writer(csvfile, **CSV_FORMAT)
        for row in rows:
            filewriter.writerow(row)
    return path


def from_csv(path):
    """Get iterable of rows from *csv_path*"""
    with Path(path).open(encoding=ENC) as csvfile:
       csvreader = csv.reader(csvfile, **CSV_FORMAT)
       for row in csvreader:
             yield row 

# Excel app
def get_app():
    try:
        return xw.apps[0]
    except:
        return xw.App()


app = get_app()    


def opener(filepath):
    app.books.api.Open(filepath, UpdateLinks=False)
    return xw.Book(filepath)

# formulas from file

def yield_refs(filepath):
    wb = opener(filepath) 
    for sheet in wb.sheets:
        print(sheet)
        try:
            df = pd.read_excel(filepath, sheetname=sheet.name) 
        except TypeError:
            print("Skipped reading", filepath, sheetname=sheet.name)
            continue
        nrows, ncols = df.shape
        for r in range(1, nrows+1+1):
            for c in range(1, ncols+1+1):
                formula = read_from_sheet(sheet, r, c)
                if has_external_ref(formula):
                    yield (sheet.name, xlref(r, c), formula)
    wb.close()


def read_from_sheet(sheet, r, c):
    if r>=256:
        return ""        
    try:
        return sheet.range(r,c).formula
    except:
        print("Failed to read cell", xlref(r,c)) 
        return ""


def yield_formulas(filepath):
   for (sheet, ref, formula) in yield_refs(filepath): 
       yield formula


def excel_files(folder=ROOT):
    return [f.__str__() for f in folder.iterdir() 
            if f.suffix in ('.xls', '.xlsx')
            and not f.name.startswith("~")]


def csv_files(folder=CSV_ROOT):
    return [f.__str__() for f in folder.iterdir() 
            if f.suffix == '.csv']


def yield_parsed(filepath, n=10):
    gen = yield_formulas(filepath)
    gen = map(get_filename, gen)
    gen = itertools.islice(gen, n)
    return gen


def dumps_links_to_csv(filepath):
    print(filepath)
    gen = yield_refs(filepath)
    csvpath = CSV_ROOT / Path(filepath).with_suffix(".csv").name 
    if not csvpath.exists():
        print("Making", csvpath)
        to_csv(gen, csvpath)        


def null(filepath):
    wb = opener(filepath) # on Windows: use raw strings to escape backslashes    
    print(wb)
    wb.close()
    

def csv_dumps():
    for filepath in excel_files():
         dumps_links_to_csv(filepath)


# to [print]
# D:\links\базовый сценарий>python excel_links.py > csv/links.txt         
def diagnose_csv():
    N = len(list(csv_files()))         
    for i, filepath in enumerate(csv_files()):
        print("\nCSV file <{}> ({}/{}):".format(filepath, i+1, N))
        gen = from_csv(filepath)
        gen = map(lambda x: get_filename(x[2]), gen)
        pick(gen)


def in_folder(path):
    root = ROOT.__str__().lower().replace("\\","/") 
    path = path.lower().replace("\\","/")
    return root in path
    
        
def pick(stream):
    fs = []
    for f in stream:
        if f not in fs:
            fs.append(f)
    fs1 = [f for f in fs if in_folder(f)]
    fs2 = [f for f in fs if f not in fs1]
    if fs1:
        print("    Links in project folder:")
        for f in sorted(fs1):
           print("    -", f)
    if fs2:       
        print("    Links outside project folder:")
        for f in sorted(fs2):
            print("    -", f) 
    if not fs:
        print("    File has no outside links.")
        

if __name__ == "__main__":
    #print(ROOT)
    
    # Job 1 - open all files
    #for f in excel_files(folder=ROOT):
    #    null (f)        
    
    # Job 2 - create CSV dumps
    csv_dumps()
    
    #Job 3 - read from CSV
    diagnose_csv()   
