# test_gs.py
from services import google_sheets as gs

if __name__ == "__main__":
    print("Sheets available (first 10 worksheet titles):")
    ss = gs._get_spreadsheet()
    print([w.title for w in ss.worksheets()][:20])
    print("\nPreview exe_summary:")
    print(gs.get_exe_summary().head())
