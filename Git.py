import psycopg2
import pandas as pd
import win32com.client

# Log-in para acceder a la database

username = "reports-user"
password = "Uiz2heeng2fite1eyaec6oe0UeSocu"

# Campos especificos para utilizar el MySQL
mydb = psycopg2.connect(
    host="172.17.4.194",
    user=username,
    password=password,
    port="5432",
    database="gitlabhq_production",
)

# Direccion del excel
pathxls = r"C:\Users\alfredo.guerra\Documents\Smartmatic book - Git.xlsx"

# Direccion de drop para los csv
path = r"C:\Users\alfredo.guerra\Documents\Export-Git"


# Funcion para refresh automatico en Excel
def excel_Git_refresh():
    xlapp = win32com.client.DispatchEx("Excel.Application")

    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(pathxls)

    # Refresh all data connections.
    wb.RefreshAll()
    wb.Save()

    # Quit
    xlapp.Quit()


def fetch_table_space1():
    myresult = pd.read_sql_query(
        "SELECT p.id project_id, p.name project_name, u.id user_id, u.username, a.access_level, a.access_name \n"
        " from users u\n"
        "Join (select user_id, source_id from members where source_type = 'Project') m  ON u.id = m.user_id \n "
        "Join projects p ON p.id = m.source_id\n"
        "join (SELECT access_level, project_id,  \n"
        "case \n"
        "when access_level = 10 then 'Guest' \n"
        "when access_level = 20 then 'Reporter'\n"
        "when access_level = 30 then 'Developer' \n"
        "when access_level = 40 then 'Master' \n"
        "when access_level = 50 then 'Administrator'\n"
        "END AS Access_name\n"
        "from project_authorizations) a\n"
        "on p.id = a.project_id\n"
        "Order by p.id;", mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_dataGit1.csv", index=False)


def fetch_table_space2():
    myresult = pd.read_sql_query(
        '''select u.id, u.name, u.username, u.email, u.state from users u
            order by u.id ;''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_Git2.csv", index=False)


def fetch_table_space3():
    myresult = pd.read_sql_query(
        '''SELECT p.id, p.name Project_Name, m.type, m.access_level, m.access_name, n.name Namespace
            from projects p
            JOIN (Select type, access_level, source_id,
	            case
	            when access_level = 10 then 'Guest' 
	            when access_level = 20 then 'Reporter'
	            when access_level = 30 then 'Developer' 
	            when access_level = 40 then 'Master' 
                when access_level = 50 then 'Administrator'
                END AS Access_name
                from members) m
	            ON m.source_id = p.id
            JOIN (select id, name from namespaces where type = 'Group') n ON p.id = n.id
            Order by p.id''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_Git3.csv", index=False)


def fetch_table_space4():
    myresult = pd.read_sql_query(
        '''SELECT u.id, u.name Username, m.type, m.access_level, m.access_name, n.name Namespace
        from users u
        JOIN (Select type, access_level, source_id,
            case
            when access_level = 10 then 'Guest' 
            when access_level = 20 then 'Reporter'
            when access_level = 30 then 'Developer' 
            when access_level = 40 then 'Master' 
            when access_level = 50 then 'Administrator'
            END AS Access_name
            from members) m
            ON u.id = m.source_id
        JOIN (select id, name from namespaces where type = 'Group') n ON u.id = n.id 
        Order by u.id''', mydb
    )

    df = pd.DataFrame(myresult)
    df.to_csv(path + "\export_Git4.csv", index=False)


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

excel_Git_refresh()
