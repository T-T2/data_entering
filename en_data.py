from db.mysqlConnection import MyPymysqlPool
from db_config import testdb_100, test74
import xlrd


def enter_data(work):
    mysql_conn = MyPymysqlPool(testdb_100)
    mysql74_conn = MyPymysqlPool(test74)
    try:
        names = work.sheet_names()
        for name in names:
        #     if name == '陈辉':
        #         continue

            _sql = """ SELECT id,`name` FROM `test`.`user`
                       WHERE `name` LIKE '%{}%'
                       AND status = 1
            """ .format(name)
            my_sql = mysql74_conn.getAll(_sql)[0]
            u_id = my_sql['id']
            _name = my_sql['name']
            table = workbook.sheet_by_name(sheet_name=name)
            table_list = table.col_values(colx=0, start_rowx=0, end_rowx=None)
            values = list()
            for v in table_list:
                mobile = int(v)
                value = (u_id, '{}'.format(_name), mobile)
                values.append(value)
            values = str(values).replace('[', '').replace(']', ';')
            # print(values)
            insert_sql = """ INSERT INTO crm.`crm_clue_export_220721`(u_id, u_name, mobile)
                       VALUES {}
            """.format(values)

            try:
                mysql_conn._cursor.execute(insert_sql)
                mysql_conn._conn.commit()
            except Exception as e:
                print(e)

    except:
        print('error')
    finally:
        mysql_conn._conn.close()
        mysql74_conn._conn.close()



if __name__ == '__main__':
    workbook = xlrd.open_workbook('data.xlsx')
    enter_data(workbook)


