import openpyxl
import os
import io
import shutil
import pyodbc
import logging

class ConnectDB:
    def __init__(self):
        self._connection = pyodbc.connect('Driver={SQL Server};Server=192.168.2.52;\
                                        Database=WebVendor_V2;uid=sa;pwd=P@ssw0rd')
        self._cursor = self._connection.cursor()

    def query(self, query):
        global result
        try:
            result = self._cursor.execute(query)
        except Exception as e:
            logging.error('error execting query "{}", error: {}'.format(query, e))
            return None
        finally:
            return result

    def update(self, sqlStatement):
        try:
            self._cursor.execute(sqlStatement)
        except Exception as e:
            logging.error('error execting Statement "{}", error: {}'.format(sqlStatement, e))
            return None
        finally:
            self._cursor.commit()

    def exec_sp(self, sqlStatement, params):
        try:
            self._cursor.execute(sqlStatement, params)
        except Exception as e:
            logging.error('error execting Statement "{}", error: {}'.format(sqlStatement, e))
            return None
        finally:
            self._cursor.commit()

    def exec_spOp(self, sqlStatement, params):
        global result
        try:
            result = self._cursor.execute(sqlStatement, params)
        except Exception as e:
            logging.error('error execting Statement "{}", error: {}'.format(sqlStatement, e))
            return None
        finally:
            return result

    def __del__(self):
        self._cursor.close()

def getDefaultParamter():
    myConnDB = ConnectDB()
    sqlStr = r'''SELECT  PARAM_VLUE
                FROM    dbo.MST_Param
                WHERE   PARAM_CODE = 'WD_DOC_GR_PATH'
                ORDER BY PARAM_SEQN'''
    result_set = myConnDB.query(sqlStr).fetchall()
    # index value
    # 0 = Result Path -> dev2\webvndGR\Result
    # 1 = Backup Result path -> dev2\webvndGR\Result_Backup
    # 2 = Log path -> dev2\log
    # 3 = IP -> \\192.168.2.52\
    src_path = str(result_set[3][0]) + str(result_set[0][0])
    des_path = str(result_set[3][0]) + str(result_set[1][0])
    log_path = str(result_set[3][0]) + str(result_set[2][0])

    # print(src_path, des_path, log_path)
    return src_path, des_path, log_path

def executeProcedure():
    myConnDB = ConnectDB()

    parm1 = 'aaa'
    params = (parm1)
    result = myConnDB.exec_spOp("""
        DECLARE @out1 varchar(255);
        DECLARE @out2 varchar(255);
        EXECUTE [dbo].[sp_kai_test] @parm1 = ?, @parm2 = @out1 OUTPUT, @parm3 = @out2 OUTPUT
        SELECT @out1, @out2
        """, params)
    # print(result)
    # print(type(result))
    for rows in result:
        print(rows[0], rows[1])
    # rows = myConnDB.fetchall()
    # while rows:
    #     print(rows)
    #     if myConnDB.nextset():
    #         rows = myConnDB.fetchall()
    #     else:
    #         rows = None

def archiveFiletoBKPath(fileFullPath, des_path):
    logging.info('Start Backup File to Destination Path [{}]'.format(fileFullPath))

    try:
        shutil.move(fileFullPath, des_path)
    except shutil.Error as err:
        logging.error('Error [{}]'.format(err))

    logging.info('End Backup File to Destination Path [{}]'.format(fileFullPath))

def main():
    executeProcedure()
    book = openpyxl.load_workbook('DailyReport.xlsx')

    sheet = book.active
    # B2|||
    # B3|||
    # B4|||
    # B5|||
    # B6|||
    # B7|||
    # B8|F8||
    # B9|F9||
    # B11|F11||
    # B13|F13||
    # B14|F14||
    # B15|F15||
    # B16|F16||
    # B17|F17||
    # B18|F18||
    # B19|D19|F19|H19
    # B20|D20|F20|H20
    # B21|D21|F21|H21
    # B23|D23|F23|H23
    # B25|D25|H25|
    # B26|D26|H26|
    # B27|D27|H27|I27
    # B28|D28|H28|
    # B29|D29|H29|
    # B30|D30|H30|
    sheet['E2'] = 'Aspire Asoke'
    sheet['B4'] = 202840
    book.save('DailyReport.xlsx')


if __name__ == '__main__':
    # Get Default Parameter from Master Parameter
    # src_path, des_path, log_path = getDefaultParamter()
    # # print(src_path, des_path, log_path)
    # # src_path = r"D:\tmp\webvndGR\Result"
    # # des_path = r"D:\tmp\webvndGR\Result_Backup"
    # logFile = log_path + '\APExcelDailyReport.log'
    #
    # logging.basicConfig(level=logging.DEBUG,
    #                     format='%(asctime)-5s [%(levelname)-8s] >> %(message)s',
    #                     datefmt='%Y-%m-%d %H:%M:%S',
    #                     filename=logFile,
    #                     filemode='a')
    #
    # logging.debug('#####################')
    # logging.info('Start Process')
    main()
    # logging.info('End Process')
    # logging.debug('#####################')
