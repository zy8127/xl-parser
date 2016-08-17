import sqlite3

from process_data.config import DATABASES
from process_data.n3_turnover import cell

conn = sqlite3.connect(database=DATABASES['default'])

cell.execute()
# print(production_info.dict)
