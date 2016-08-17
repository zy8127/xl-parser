from win32com.client import Dispatch
import os

from pandas import DataFrame, to_datetime, concat
from calendar import monthrange
from datetime import timedelta, date

from ..config import DIRS, WORKBOOKS, CONN, SD


class ExcelBook():
    def __init__(self, filename):
        self.filename = os.path.join(DIRS['SOURCE_DIR'], WORKBOOKS[filename])
        self.connection = Dispatch('ADODB.Connection')
        self.connection.Open(
            '''
            Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source={0};
            Extended Properties="Excel 12.0 Xml;
            HDR=YES;
            IMEX=1";
            '''.format(self.filename)
        )

    def sheets(self):
        result = []
        recordset = self.connection.OpenSchema(20)
        while not recordset.EOF:
            result.append(recordset.Fields[2].Value)
            recordset.MoveNext()
        recordset.Close()
        del recordset
        return result

    def sheet(self, name, sql, stack=False, icol=False):
        """icol: insert column"""
        s = ExcelSheet(self.connection, name, sql=sql, stack=stack, icol=icol)
        s.data()
        s.to_db()
        return s.df

    def close(self):
        try:
            self.connection.Close()
            del self.connection
        except:
            pass

    def __del__(self):
        self.close()

    def __enter__(self):
        print('---connect to {0}'.format(self.filename))
        return self

    def __exit__(self, exc_type, exc_value, trackback):
        self.close()


class ExcelSheet():
    def __init__(self,
                 connection,
                 sheet,
                 sql,
                 stack=False,
                 icol=False):

        self.sheet = sheet
        self.connection = connection
        self.recordset = Dispatch('ADODB.Recordset')
        self.recordset.Open(sql, self.connection, 0, 1)
        self.stack = stack
        # self.sd = SD
        self.icol = icol
        self.df = None
        print('  |--connect to {0}'.format(self.sheet))

    def column_dates(self):
        # current_month_days
        cmd = monthrange(
            SD.year,
            SD.month + 1
        )
        # current_month_last_day
        cmld = date(
            SD.year,
            SD.month + 1,
            cmd[1]
        )
        return [SD + timedelta(days=i) for i in range(cmd[1])] + \
            [cmld - timedelta(days=1) for j in
             range((len(self.recordset.Fields)-cmd[1]-3))]

    def column_names(self):
        if self.stack:
            return ['产品图号', '产品类别'] + \
                self.column_dates() + ['合计']
        else:
            return [field.Name for field in self.recordset.Fields]

    def data(self):
        self.df = DataFrame(data=list(self.recordset.GetRows()))
        self.df = self.df.T
        self.df.columns = self.column_names()
        self.df = self.df.set_index('产品图号')

        if self.sheet == '东海.外协':
            col_names = ['产品类别', '外协盘存', '本月出库',
                         '本月入库', '本月结存'] + \
                         [SD + timedelta(days=i)
                          for i in range(31)] + ['合计'] + \
                         [SD + timedelta(days=i)
                          for i in range(31)] + ['合计1']
            self.df.columns = col_names
            del self.df['产品类别']
            dd = self.df.iloc[:, 0:1]
            # dd.columns = ['外协盘存']
            dd = dd[dd['外协盘存'] > 0]
            dd['外协盘存'] = dd['外协盘存'].astype(int)
            dd.to_sql(name='东海.外协盘点', con=CONN,
                      flavor='sqlite', if_exists='replace')

            del self.df['合计']
            del self.df['合计1']
            del self.df['外协盘存']
            del self.df['本月出库']
            del self.df['本月入库']
            del self.df['本月结存']

            df1 = self.df.iloc[:, 0:31]
            df1 = df1.stack()
            df1.index.names = ['产品图号', '日期']
            df1 = df1.to_frame()
            df1.columns = ['数量']
            df1.insert(loc=1, column='工序', value='W1TO')

            df2 = self.df.iloc[:, 0:31]
            df2 = df2.stack()
            df2.index.names = ['产品图号', '日期']
            df2 = df2.to_frame()
            df2.columns = ['数量']
            df2.insert(loc=1, column='工序', value='WOT1')

            self.df = DataFrame(concat([df1, df2]))
            self.df['数量'] = self.df['数量'].astype(int)
            self.df = self.df[self.df['数量'] != 0]

        self.df = self.conditions()

        return self.df

    def conditions(self):
        if self.stack:
            for cn in self.df.columns:
                if cn in ['直径', '长度', '成品长度', '产品类别', '合计']:
                    del self.df[cn]
            self.df = self.df.stack()
            self.df.index.names = ['产品图号', '日期']
            self.df = self.df.to_frame()
            self.df.columns = ['数量']
            self.df['数量'] = self.df['数量'].astype(int)
            self.df = self.df[self.df['数量'] != 0]

        if self.icol:
            self.df.insert(loc=1, column='工序', value=self.icol)

        for cn in self.df.columns:
            if cn in ['W3盘存', 'W2盘存', 'W4盘存']:
                self.df[cn] = self.df[cn].astype(int)

        if self.sheet in ['一部.工序监控', 'G加.工序数据']:
            self.df['日期'] = self.df['日期'].apply(convert_data)
            self.df['日期'] = self.df['日期'].dt.date
            self.df['数量'] = self.df['数量'].astype(int)

        if self.sheet == 'G加.工序数据':
            self.df['工序'] = self.df['工序'].replace(
                ['粗磨', '淬火', '回火', '半中磨', '精车',
                 '镀前磨削', '电镀', '镀后', 'GP12检验'],
                ['W4粗磨', 'W4淬火', 'W4回火', 'W4中磨', 'W4精车',
                 'W4镀前', 'W4电镀', 'W4镀后', 'W4GP12'])
            self.df = self.df[self.df['工序'].str.startswith('W4')]

        return self.df

    def to_db(self):
        self.df.to_sql(name=self.sheet, con=CONN,
                       flavor='sqlite', if_exists='replace')

    def close(self):
        try:
            self.recordset.Close()
            del self.recordset
        except:
            pass

    def __del__(self):
        # self.close()
        pass


def convert_data(s):
    d = str(s)
    return to_datetime(d)
