import pytest
from fixture.application import Application
from comtypes.client import CreateObject
import os


@pytest.fixture(scope="session")
def app(request):
    fixture = Application("C:\\ToolsNatasha\\FreeAddressBookPortable\\AddressBook.exe")
    request.addfinalizer(fixture.destroy)
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