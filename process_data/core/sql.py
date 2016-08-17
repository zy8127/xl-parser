import os


def get_sql(fn):
    f = open(fn, encoding='utf-8')
    c = [l.strip() for l in f.readlines()]
    sql = ' '.join(c)
    return sql


def generate_sql_dict(path):
    return {
        file.name.strip('.sql'): get_sql(os.path.join(path, file.name))
        for file in os.scandir(path)
        if '.sql' in file.name
        }
