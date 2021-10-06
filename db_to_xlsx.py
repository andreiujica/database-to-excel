from sqlalchemy import create_engine, MetaData, text
from sqlalchemy.engine.base import Engine
import xlsxwriter
import argparse
import logging
import os


def initialize_logs() -> None:
    """ Establish minimum logging level, logging format and log file name """

    logging.basicConfig(level=logging.INFO, format="%(asctime)s :: %(levelname)s :: %(funcName)s:: %(message)s", filename="db_to_xlsx.log")

def get_input_parameters() -> argparse.Namespace:
    """ Use argparse library to restrict input data, display information about 
        input fields and altogether make script more command-line friendly """

    parser = argparse.ArgumentParser(description="This Python script converts a Database into a Microsoft Excel spreadsheet")

    parser.add_argument("-f","--flavour", choices=["sqlite", "mysql", "postgresql", "oracle", "mssql"], 
                        required = True, help = "Flavour of the database to be converted")
    parser.add_argument("-u", "--username", required= False, help="Username for connecting to database")
    parser.add_argument("-pw","--password", required=False, help = "Password for connecting to the database")
    parser.add_argument("-hn", "--hostname", required = False, help = "IP of the database server")
    parser.add_argument("-p", "--port", required=False, help="Database port on the server")
    parser.add_argument("-db", "--database", required = True, help = "Name of the database (for SQLite add .db extension)")

    args = parser.parse_args()
    logging.info("Input arguments are ok.")
    return args

def create_db_connection(dialect, database, username , password , host , port) -> Engine:
    """ Use input data to create SQLalchemy engine - critical for connecting to the database, 
        depending on db flavour and log messages"""

    # Dictionary with database drivers used in SQLalchemy URL generation
    drivers = {"mysql" : "+pymysql", "postgresql" : "+psycopg2", "oracle" : "", "mssql" : "+pymssql"}

    if dialect == "sqlite" :
        # Only if sqlite database exists in directory try to open it, otherwise
        # create_engine creates a new ".db" file name {database}.db
        if os.path.isfile(f"./{database}"):
            engine = create_engine(f'{dialect}:///{database}', echo = False)
            logging.info("Opened SQLite Database.")
            print("Connected to SQLite Database!")
        else:
            print("Database does not exist")
            logging.error("SQLite file does not exist in the current directory.")
            exit()
    else:
        try:
            print("Connecting to Database server...")
            engine = create_engine(f"{dialect}{drivers[dialect]}://{username}:{password}@{host}:{port}/{database}", echo = False)
            meta = MetaData()
            meta.reflect(bind=engine)
            print(f"Connected to {dialect} server!")
            logging.info(f"Connected to {dialect} server on {host}")
        except:
            print(f"Can't connect to {dialect} server on {host}")
            logging.exception(f"Could not connect to {dialect} server on {host}")
            exit()
    
    return engine

def get_column_names(engine, table):
    """Self-explanatory function, gets column names from engine metadata"""

    meta = MetaData()
    meta.reflect(bind=engine)
    return list(meta.tables[table].columns.keys())

def get_table_name(sel: str) -> str:
    """Helper function to get table name from SQL SELECT statement"""

    # Split selection by delimiter "FROM" and get the first element of that list
    try:
        table = sel.split("from")[1].split()[0]
    except:
        table = sel.split("FROM")[1].split()[0]

    return table

def get_data(engine):
    """Gets SELECT statement from user, asserts if valid then executes SQL statement and fetches data"""

    while True:
        try: 
            sel = input("Enter Query for any table here (only SELECT permited): ")
            table = get_table_name(sel)
            
            # Check if this is indeed a SELECT statement
            assert sel.upper().startswith("SELECT")
            break

        except AssertionError:
            print("Not a SELECT Query. Only SELECT Queries are permitted.")
            logging.warning("A Query other than SELECT has been requested")
            continue
        except KeyboardInterrupt:
            exit()

    # Connection needed in order to use "execute" method
    conn = engine.connect()

    try:
        query = conn.execute(text(sel))
        logging.info("Fetched Database data.")
        return query, table
    except:
        print("Not a valid SQL Query")
        logging.error("Invalid SELECT statement requested.")
        exit()
    

def write_excel_sheet(workbook, worksheet, data, headings):
    """ Helper function for writing Excel file """

    # Name of the worksheet will be the name of the table
    worksheet = workbook.add_worksheet(worksheet)

    # Write headings on the first line
    col = 0
    for _ in range(len(headings)):
        worksheet.write(0, col , headings[col])
        col += 1

    # Write the data, row by row
    row, col = 1,0
    for entry in data:
        worksheet.write_row(row, col, entry)
        row += 1

def write_excel_file(engine: Engine, flavour: str, database: str) -> None:
    """ Main Xlsx writing function, gets the data, creates workbook and worksheet, populates worksheet, logs everything """

    # Fetch data that will populate excel sheets
    data, table = get_data(engine)

    # As the database name for sqlite contains ".db" we need to erase that extension
    if flavour == "sqlite":
        workbook = xlsxwriter.Workbook(f"./excel_files/{database[:len(database) - 3]}.xlsx")
    else:
        workbook = xlsxwriter.Workbook(f"./excel_files/{database}.xlsx")
    
    write_excel_sheet(workbook, table, data, get_column_names(engine, table))

    print("Success! Excel files have been written.")
    logging.info("Excel files have been written succesfully.")
    workbook.close()

def main(): 
    initialize_logs()

    args = get_input_parameters()

    engine = create_db_connection(args.flavour, args.database, args.username, args.password, args.hostname, args.port)

    write_excel_file(engine, args.flavour, args.database)
    

if __name__ == "__main__":
    main()