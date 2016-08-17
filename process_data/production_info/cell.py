import os

from ..core.connection import ExcelBook
from ..core.sql import generate_sql_dict
from ..config import DIRS, WORKBOOKS, DATABASES

path = os.path.dirname(__file__)
ws = generate_sql_dict(path=path)


def execute():
    b = ExcelBook(filename='产品信息')
    a = b.sheet(name='产品信息', sql=ws['产品信息'], stack=False, icol=False)
    print(a)
