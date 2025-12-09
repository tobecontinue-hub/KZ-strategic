
import pandas as pd

class DataLoader:
    def __init__(self, excel_path):
        self.path = excel_path
        self._cache = {}

    def list_sheets(self):
        xl = pd.ExcelFile(self.path)
        return xl.sheet_names

    def get_sheet_df(self, sheet_name):
        # cache simple
        if sheet_name in self._cache:
            return self._cache[sheet_name]
        xl = pd.ExcelFile(self.path)
        if sheet_name not in xl.sheet_names:
            # try lower-case match
            names = {n.lower(): n for n in xl.sheet_names}
            if sheet_name.lower() in names:
                sheet_name = names[sheet_name.lower()]
            else:
                raise ValueError(f"Sheet {sheet_name} not found. Available: {xl.sheet_names}")
        df = pd.read_excel(xl, sheet_name=sheet_name)
        self._cache[sheet_name] = df
        return df
