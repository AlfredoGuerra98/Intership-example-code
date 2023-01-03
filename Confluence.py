import mysql.connector
import pandas as pd
from datetime import datetime
import win32com.client

# username y password de MySql 
# Log-in para acceder a la database

username = "alfredo.guerra"
password = "WUDJmV8a3L"

# Campos especificos para utilizar el MySQL
mydb = mysql.connector.connect(
    host="10.4.82.71",
    user=username,
    password=password,
    database="confluence_5912"
)

# path siendo en donde se va depositar los csv, no remuevan la r antes del ""
# Direccion del excel
pathxls = r"C:\Users\alfredo.guerra\Documents\Smartmatic book - Confluence.xlsx"

# Direccion de drop para los csv
path = r"C:\Users\alfredo.guerra\Documents\Export-Confluence"


# Funcion para refresh automatico en Excel
def excel_refresh_confluence():
    xlapp = win32com.client.DispatchEx("Excel.Application")

    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(pathxls)

    # Refresh all data connections.
    wb.RefreshAll()
    wb.Save()

    # Quit
    xlapp.Quit()


def fetch_table_space1():
    myresult2 = "SELECT sp.LASTMODDATE, s.SPACEKEY, s.SPACENAME, sp.PERMGROUPNAME, u.user_name, sp.PERMTYPE " \
                "FROM SPACES s JOIN SPACEPERMISSIONS sp ON s.SPACEID = sp.SPACEID " \
                "JOIN cwd_group g ON sp.PERMGROUPNAME = g.group_name JOIN cwd_membership m ON g.id = m.parent_id " \
                "JOIN cwd_user u ON m.child_user_id = u.id WHERE s.SPACEKEY not like '~%'  AND  sp.LASTMODDATE >= "

    year = 2000

    while True:
        temp = myresult2 + " '" + str(year) + "-01-01' AND sp.LASTMODDATE <= '" + str(year) + "-12-31'"
        myresult = pd.read_sql_query(
            temp, mydb
        )

        if year == 2000:
            myresult.to_csv(path + "\export_data1.csv", index=False)
        else:
            myresult.to_csv(path + "\export_data1.csv", index=False, mode='a', header=False)

        year += 1
        if year > datetime.now().year:
            break


def fetch_table_space2():
    myresult2 = "SELECT p.LASTMODDATE, s.SPACEKEY, s.SPACENAME, c.user_name, p.PERMTYPE FROM SPACES s " \
                "JOIN SPACEPERMISSIONS p ON s.SPACEID = p.SPACEID JOIN user_mapping u ON p.PERMUSERNAME = u.user_key " \
                "JOIN cwd_user c ON u.username = c.user_name " \
                "WHERE s.SPACEKEY not like '~%' AND p.LASTMODDATE >= "
    year = 2000

    while True:
        temp = myresult2 + " '" + str(year) + "-01-01' AND p.LASTMODDATE <= '" + str(year) + "-12-31'"
        myresult = pd.read_sql_query(
            temp, mydb
        )

        if year == 2000:
            myresult.to_csv(path + "\export_data2.csv", index=False)
        else:
            myresult.to_csv(path + "\export_data2.csv", index=False, mode='a', header=False)

        year += 1
        if year > datetime.now().year:
            break


def fetch_table_space3():
    myresult2 = "SELECT sp.LASTMODDATE, s.SPACEKEY, s.SPACENAME, sp.PERMGROUPNAME, sp.PERMTYPE " \
                "FROM SPACES s JOIN SPACEPERMISSIONS sp ON s.SPACEID = s.SPACEID " \
                "WHERE sp.PERMTYPE = 'SETSPACEPERMISSIONS' AND sp.PERMGROUPNAME != '' AND  sp.LASTMODDATE >= "
    year = 2000

    while True:
        temp = myresult2 + " '" + str(year) + "-01-01' AND sp.LASTMODDATE <= '" + str(year) + "-12-31'"
        myresult = pd.read_sql_query(
            temp, mydb
        )

        if year == 2000:
            myresult.to_csv(path + "\export_data3.csv", index=False)
        else:
            myresult.to_csv(path + "\export_data3.csv", index=False, mode='a', header=False)

        year += 1
        if year > datetime.now().year:
            break


def fetch_table_space4():
    myresult = pd.read_sql_query(
        '''SELECT u.lower_user_name, d.directory_name FROM cwd_user u 
        JOIN cwd_membership m ON u.id = child_user_id JOIN cwd_group g ON m.parent_id = g.id 
        JOIN SPACEPERMISSIONS sp ON g.group_name = sp.PERMGROUPNAME JOIN cwd_directory d on u.directory_id = d.id 
        WHERE u.active = 'T' AND d.active = 'T' GROUP BY u.lower_user_name, d.directory_name 
        ORDER BY d.directory_name;''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_data4.csv", index=False)


def fetch_table_space5():
    print("Table 1")
    fetch_table_space1()
    print("Table 2")
    fetch_table_space2()
    print("Table 3")
    fetch_table_space3()
    print("Table 4")
    fetch_table_space4()


# Inicio del programa
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

excel_refresh_confluence()
