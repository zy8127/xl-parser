import os

from ..core.connection import ExcelBook
from ..core.sql import generate_sql_dict
from ..config import DIRS, WORKBOOKS, DATABASES


def execute():
    path = os.path.dirname(__file__)
    ws = generate_sql_dict(path=path)

    with ExcelBook(filename='拉拔转序') as b:
        b.sheet('tc拉拔转序', ws['拉拔转序'], stack=True, icol=False)
        b.sheet('ntc拉拔盘点', ws['拉拔汇总'], stack=False, icol=False)
