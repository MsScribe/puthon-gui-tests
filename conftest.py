import pytest
from fixture.application import Application
from comtypes.client import CreateObject
import os
import json
import os.path


fixture = None
target = None


def load_config(file):
    global target
    if target is None:
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), file)
        with open(config_file) as f:
            target = json.load(f)
    return target


@pytest.fixture(scope="session")
def app(request):
    global fixture
    app_path_config = load_config(request.config.getoption("--target"))["gui"]
    if fixture is None or not fixture.is_valid():
        fixture = Application(target=app_path_config["baseURL"])
    def fin():
        fixture.destroy()
    request.addfinalizer(fin)
    return fixture


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("data_"):
            xl = CreateObject("Excel.Application")
            wb = xl.Application.Workbooks.Open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data/%s.xlsx" % fixture[5:]))
            worksheet = wb.Sheets[1]
            testdata = []
            for row in range(1, 11):
                data = worksheet.Cells[row, 1].Value()
                testdata.append(data)
            xl.Quit()
            metafunc.parametrize(fixture, testdata, ids=[str(x) for x in testdata])


def pytest_addoption(parser):
    parser.addoption("--target", action="store", default="target.json")