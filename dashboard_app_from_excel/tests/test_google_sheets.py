import importlib
import sys
import types
import pandas as pd
import pytest


class FakeWorksheet:
    def __init__(self, title, records):
        self.title = title
        self._records = records

    def get_all_records(self):
        return self._records


class FakeSpreadsheet:
    def __init__(self):
        self.last_worksheet_name = None
        self.by_index_calls = []
        self.ws_by_name = {}
        self.ws_by_index = {}
        self.opened_key = None

    def worksheet(self, name):
        self.last_worksheet_name = name
        if name in self.ws_by_name:
            return self.ws_by_name[name]
        raise Exception("Worksheet not found")

    def get_worksheet(self, index):
        self.by_index_calls.append(index)
        if index in self.ws_by_index:
            return self.ws_by_index[index]
        # Default empty worksheet for any index
        return FakeWorksheet(f"idx_{index}", [])

    def worksheets(self):
        # Optional: used by manual test script
        return list(self.ws_by_name.values()) + list(self.ws_by_index.values())


class FakeClient:
    def __init__(self, spreadsheet: FakeSpreadsheet):
        self._spreadsheet = spreadsheet

    def open_by_key(self, key):
        self._spreadsheet.opened_key = key
        return self._spreadsheet


class FakeCredentials:
    @classmethod
    def from_service_account_file(cls, filename, scopes=None):
        return object()


@pytest.fixture
def gs_module(monkeypatch):
    """
    Prepare a clean import of services.google_sheets with external deps faked in sys.modules
    to avoid importing real gspread/google libraries.
    """
    # Ensure a clean import each time
    sys.modules.pop('services.google_sheets', None)

    # Build fake gspread module
    gspread_mod = types.ModuleType('gspread')

    def _not_configured(*args, **kwargs):
        raise RuntimeError("Test must monkeypatch gspread.authorize in each test")

    gspread_mod.authorize = _not_configured

    # Build fake google.oauth2.service_account with Credentials
    google_mod = types.ModuleType('google')
    oauth2_mod = types.ModuleType('google.oauth2')
    service_account_mod = types.ModuleType('google.oauth2.service_account')
    service_account_mod.Credentials = FakeCredentials

    sys.modules['gspread'] = gspread_mod
    sys.modules['google'] = google_mod
    sys.modules['google.oauth2'] = oauth2_mod
    sys.modules['google.oauth2.service_account'] = service_account_mod

    # Import target module
    gs = importlib.import_module('services.google_sheets')

    # Normalize config and reset cache for test isolation
    monkeypatch.setattr(gs, 'GOOGLE_SHEET_ID', 'TEST_SHEET_ID', raising=False)
    monkeypatch.setattr(gs, 'SERVICE_ACCOUNT_FILE', 'test_service_account.json', raising=False)
    monkeypatch.setattr(gs, '_cache', {"ts": 0, "spreadsheet": None}, raising=False)

    return gs


@pytest.fixture
def fake_spreadsheet():
    return FakeSpreadsheet()


@pytest.fixture(autouse=True)
def reset_cache_between_tests(gs_module, monkeypatch):
    # Ensure each test starts with a clean cache
    monkeypatch.setattr(gs_module, '_cache', {"ts": 0, "spreadsheet": None}, raising=False)


def test_missing_service_account_raises_file_not_found(gs_module, monkeypatch):
    # Make service account file appear missing
    monkeypatch.setattr(gs_module.os.path, 'exists', lambda p: False)

    with pytest.raises(FileNotFoundError):
        # Calls _get_client -> checks file -> raises
        gs_module._get_spreadsheet()


def test_open_by_key_called_and_dataframe_returned(gs_module, monkeypatch, fake_spreadsheet):
    # Service account present
    monkeypatch.setattr(gs_module.os.path, 'exists', lambda p: True)

    # Setup spreadsheet to return a named worksheet
    records = [{"a": 1, "b": 2}, {"a": 3, "b": 4}]
    fake_spreadsheet.ws_by_name['exe_summary'] = FakeWorksheet('exe_summary', records)

    # gspread.authorize -> FakeClient(fake_spreadsheet)
    def _auth(_):
        return FakeClient(fake_spreadsheet)

    monkeypatch.setattr(gs_module.gspread, 'authorize', _auth)

    # Call high-level helper
    df = gs_module.get_exe_summary()

    # Verify path through gspread client and DataFrame content
    assert fake_spreadsheet.opened_key == 'TEST_SHEET_ID'
    assert fake_spreadsheet.last_worksheet_name == 'exe_summary'
    assert list(df.columns) == ['a', 'b']
    assert len(df) == 2


def test_caching_within_ttl_and_refresh_after_ttl(gs_module, monkeypatch):
    # Service account present
    monkeypatch.setattr(gs_module.os.path, 'exists', lambda p: True)

    # Control time
    class Clock:
        def __init__(self, t):
            self.t = t
        def time(self):
            return self.t
        def advance(self, dt):
            self.t += dt

    clock = Clock(1_000.0)
    monkeypatch.setattr(gs_module.time, 'time', clock.time)

    # Incrementing clients to show refresh
    created = []

    def make_client():
        sp = FakeSpreadsheet()
        created.append(sp)
        return FakeClient(sp)

    def _auth(_):
        return make_client()

    monkeypatch.setattr(gs_module.gspread, 'authorize', _auth)

    # Short TTL
    monkeypatch.setattr(gs_module, 'SHEET_CACHE_TTL', 10, raising=False)

    # First fetch -> creates client/spreadsheet[0]
    ss1 = gs_module._get_spreadsheet()
    assert ss1 is created[0]

    # Within TTL -> returns cached
    clock.advance(5)
    ss2 = gs_module._get_spreadsheet()
    assert ss2 is ss1

    # After TTL -> refresh to a new spreadsheet
    clock.advance(6)
    ss3 = gs_module._get_spreadsheet()
    assert ss3 is not ss1
    assert ss3 is created[1]

    # Force refresh regardless of TTL
    clock.advance(1)
    ss4 = gs_module._get_spreadsheet(force_refresh=True)
    assert ss4 is not ss3
    assert ss4 is created[2]


def test_sheet_to_df_falls_back_to_index_when_name_missing(gs_module, monkeypatch, fake_spreadsheet):
    monkeypatch.setattr(gs_module.os.path, 'exists', lambda p: True)

    # No sheet by name, but index 3 exists with records
    records = [{"x": 10}, {"x": 20}, {"x": 30}]
    fake_spreadsheet.ws_by_index[3] = FakeWorksheet('idx_3', records)

    def _auth(_):
        return FakeClient(fake_spreadsheet)

    monkeypatch.setattr(gs_module.gspread, 'authorize', _auth)

    df = gs_module.sheet_to_df('missing_name', worksheet_index=3)

    assert 3 in fake_spreadsheet.by_index_calls
    assert list(df.columns) == ['x']
    assert len(df) == 3


@pytest.mark.parametrize(
    'helper_name, expected_sheet', [
        ('get_brand_promise', 'brand_promise'),
        ('get_value_map', 'value_map'),
        ('get_key_decisions', 'key_desicions'),  # note: module currently uses this spelling
    ]
)
def test_helpers_request_expected_sheet_names(gs_module, monkeypatch, fake_spreadsheet, helper_name, expected_sheet):
    monkeypatch.setattr(gs_module.os.path, 'exists', lambda p: True)

    # Provide worksheets for names under test
    fake_spreadsheet.ws_by_name[expected_sheet] = FakeWorksheet(expected_sheet, [{"ok": 1}])

    def _auth(_):
        return FakeClient(fake_spreadsheet)

    monkeypatch.setattr(gs_module.gspread, 'authorize', _auth)

    # Call corresponding helper
    helper = getattr(gs_module, helper_name)
    df = helper()

    assert fake_spreadsheet.last_worksheet_name == expected_sheet
    assert list(df.columns) == ['ok']
    assert len(df) == 1
