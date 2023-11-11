import glob
from pathlib import Path
import pandas as pd

filepaths = glob.glob(".\\invoices\\*.xlsx")

for filepath in filepaths:
    invoice_num = Path(filepath.stem)
    df = pd.read_excel(Path(filepath))
