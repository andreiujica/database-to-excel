from sqlalchemy import create_engine, MetaData, text
import pandas as pd
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

def create_db_connection(dialect, database, username , password , host , port):
    """ Use input data to create SQLalchemy engine - critical for connecting to the database, 
        depending on db flavour and log messages"""

    # Dictionary with database drivers used in SQLalchemy URL generation
    drivers = {"mysql" : "+pymysql", "postgresql" : "+psycopg2", "oracle" : "+cx_oracle", "mssql" : "+pymssql"}

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

def get_column_names(result):
    """Self explanatory function, get column names from cursor result keys description"""

    return list(result.keys())

def get_data(engine):
    """Gets SELECT statement from user, asserts if valid then executes SQL statement and fetches data"""

    while True:
        try: 
            sel = input("Enter Query for any table here (only SELECT permited): ")
            
            # Check if this is indeed a SELECT statement
            assert sel.upper().find("SELECT")>-1 and sel.upper().find("FROM")>-1
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
        result = conn.execute(text(sel))
        logging.info("Fetched Database data.")
        return result
    except:
        print("Not a valid SQL Query")
        logging.error("Invalid SELECT statement requested.")
        exit()
    
def write_excel_file(engine, flavour: str, database: str) -> None:
    """ Main Xlsx writing function, gets the data, creates workbook and worksheet, populates worksheet, logs everything """

    # Fetch data that will populate excel sheets
    data = get_data(engine)

    # Create Pandas DataFrame using from_records method
    df = pd.DataFrame.from_records(data, columns=get_column_names(data))
    

    # As the database name for sqlite contains ".db" we need to erase that extension
    if flavour == "sqlite":
        writer = pd.ExcelWriter(f"./excel_files/{database[:len(database) - 3]}.xlsx", engine = "xlsxwriter")
    else:
        writer = pd.ExcelWriter(f"./excel_files/{database}.xlsx", engine = "xlsxwriter")
    
    df.to_excel(writer, sheet_name = "Query Result", index = False)

    print("Success! Excel files have been written.")
    logging.info("Excel files have been written succesfully.")
    
    writer.save()


def main(): 
    initialize_logs()

    args = get_input_parameters()

    engine = create_db_connection(args.flavour, args.database, args.username, args.password, args.hostname, args.port)

    write_excel_file(engine, args.flavour, args.database)
    

if __name__ == "__main__":
    main()
