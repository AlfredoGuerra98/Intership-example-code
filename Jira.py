import mysql.connector
import pandas as pd
import win32com.client

# Log-in para acceder a la database

username = "alfredo.guerra"
password = "WUDJmV8a3L"

# Campos especificos para utilizar el MySQL
mydb = mysql.connector.connect(
    host="10.4.82.60",
    user=username,
    password=password,
    database="jiradb"
)

# Direccion del excel
pathxls = r"C:\Users\alfredo.guerra\Documents\Smartmatic book - Jira.xlsx"

# Direccion de drop para los csv
path = r"C:\Users\alfredo.guerra\Documents\Export - Jira"


# Funcion para refresh automatico en Excel
def excel_jira_refresh():
    xlapp = win32com.client.DispatchEx("Excel.Application")

    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(pathxls)

    # Refresh all data connections.
    wb.RefreshAll()
    wb.Save()

    # Quit
    xlapp.Quit()


def fetch_table_space1():
    print("Table 1")
    myresult = pd.read_sql_query(
        '''SELECT DISTINCT p.pkey, p.pname, r.ROLETYPEPARAMETER GRUPO, rr.NAME ROL, s.PERMISSION_KEY, u.user_name
       FROM projectroleactor r
       left outer join projectrole rr on rr.id = r.projectroleID
       INNER JOIN schemepermissions s ON s.perm_parameter = r.roletypeparameter
       INNER JOIN project p on p.id = r.pid
       JOIN cwd_group g ON r.ROLETYPEPARAMETER = g.group_name
       JOIN cwd_membership m ON g.id = m.parent_id
       INNER JOIN cwd_user u ON m.child_id = u.id''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_dataJira1.csv", index=False)


def fetch_table_space2():
    print("Table 2")
    myresult = pd.read_sql_query(
        '''SELECT p.pkey, s.PERMISSION_KEY, p.pname, u.user_name, rr.NAME ROL
       FROM projectroleactor r
       JOIN projectrole rr ON rr.id = r.projectroleID
       JOIN project p on p.id = r.pid
       JOIN schemepermissions s ON s.perm_parameter = r.roletypeparameter
       JOIN cwd_group g ON r.ROLETYPEPARAMETER = g.group_name
       JOIN cwd_membership m ON g.id = m.parent_id
       JOIN cwd_user u ON m.child_id = u.id''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_dataJira2.csv", index=False)


def fetch_table_space3():
    print("Table 3")
    myresult = pd.read_sql_query(
        '''SELECT p.pkey, s.PERMISSION_KEY, p.pname, r.ROLETYPEPARAMETER GRUPO, rr.NAME ROL
       FROM projectroleactor r
       JOIN projectrole rr ON rr.id = r.projectroleID
       JOIN project p on p.id = r.pid
       JOIN schemepermissions s ON s.perm_parameter = r.roletypeparameter''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_dataJira3.csv", index=False)


def fetch_table_space4():
    print("Table 4")
    myresult = pd.read_sql_query(
        '''select u.lower_user_name, d.directory_name
            FROM cwd_user u
            INNER JOIN cwd_membership m ON u.id = m.child_id
            INNER JOIN cwd_directory d on u.directory_id = d.id
            WHERE u.id in (
                select max(u.id) FROM cwd_user u
                INNER JOIN cwd_directory d on u.directory_id = d.id
                WHERE u.active = 1
                GROUP BY u.lower_user_name)
            AND m.lower_parent_name in ('jira-users', 'gr-jira', 'jira-administrators') ;''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_dataJira4.csv", index=False)


def fetch_table_space5():
    print("Table 1")
    fetch_table_space1()
    print("Table 2")
    fetch_table_space2()
    print("Table 3")
    fetch_table_space3()
    print("Table 4")
    fetch_table_space4()


print("Escoja una table para desplegar en Excel (#1 al 4)")
print("Si desea desplegar todas las tablas, seleccione 5.")
options = int(input("Escriba su seleccion: "))

print("---------------------------------------------------------------------")

while options != 0:
    if options == 1:
        print("Table 1")
        fetch_table_space1()
        break

    elif options == 2:
        print("Table 2")
        fetch_table_space2()
        break
    elif options == 3:
        print("Table 3")
        fetch_table_space3()
        break

    elif options == 4:
        print("Table 4")
        fetch_table_space4()
        break

    elif options == 5:
        fetch_table_space5()
        break

    else:
        print("Opcion no valida")
        break

excel_jira_refresh()
