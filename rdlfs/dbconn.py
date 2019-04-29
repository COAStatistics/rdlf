import pyodbc
import sys
from log import err_log


class DatabaseConnection:
   
    DISASTER =\
    """
    SELECT [evt_name], [approveCrop], [apr_area], [sbdy_amt]
    FROM [disaster].[dbo].[107acdList_farmerSurvey]
    WHERE [appID] = convert(nvarchar(255), ?)
    """
    
    CROP_SUBSIDY =\
    """
    SELECT [crop], [price], [period]
    FROM [fallow].[dbo].[107transferSubsidy_farmerSurvey]
    WHERE [id] = convert(nvarchar(255), ?)
    """
    
    __pid = None
    args = 'Driver={ODBC Driver 13 for SQL Server};Server=172.16.21.8;Database=%s;Trusted_Connection=yes;'
    __instance = None
    
    def __init__(self, db_name='fallow'):
        self.conn = pyodbc.connect(DatabaseConnection.args % db_name)
        self.cur = self.conn.cursor()
            
    def get_disaster(self) -> list:
        d_l = []
        try:
            self.cur.execute(DatabaseConnection.DISASTER, DatabaseConnection.__pid)
        except Exception:
            info = sys.exc_info()
            err_log.error(info[0], '\t', info[1])
        else:
            rows = self.cur.fetchall()
            if rows:
                for i in rows:
                    l = [i.evt_name, i.approveCrop, str(round(i.apr_area, 4)), str(int(i.sbdy_amt))]
                    try:
                        assert len(l[0]) != 0 and len(l[1]) !=0 and float(l[2]) > 0 and int(l[3]) > 0
                    except AssertionError:
                        err_log.error('AssertionError: ', self.get_disaster.__name__, ' id=', DatabaseConnection.pid, ' ', l)
                    else:
                        d_l.append(l)
        return d_l
    
    def get_crop_subsidy(self) -> list:
        c_s_l = []
        try: 
            self.cur.execute(DatabaseConnection.CROP_SUBSIDY, DatabaseConnection.__pid)
        except Exception:
            info = sys.exc_info()
            err_log.error(info[0], '\t', info[1])
        else:
            rows = self.cur.fetchall()             
            if rows:
                for record in rows:
                    l = list(record)
                    try:
                        assert float(l[1]) > 0 and l[2] == '1'
                    except AssertionError:
                        err_log.error('AssertionError: ', self.get_crop_subsidy.__name__, ' id=', DatabaseConnection.__pid, ' ', l)
                    else:
                        c_s_l.append(l)
        return c_s_l

    def close_conn(self) -> None:
        self.cur.close()
        self.conn.close()
    
    @classmethod
    def set_pid(cls, pid):
        cls.__pid = pid

    @staticmethod
    def get_db_instance():
        if DatabaseConnection.__instance is not None:
            return DatabaseConnection.__instance
        else:
            DatabaseConnection.__instance = DatabaseConnection()
            return DatabaseConnection.__instance
