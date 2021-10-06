<div id="top"></div>

<!-- SCRIPT INFO -->
<br />
<div align="center">
  <h3 align="center">Database-To-Excel</h3>
  <p align="center">
    A Short Python script that converts a database to an excel sheet
  </p>
</div>


<!-- ABOUT THE SCRIPT -->
## About The Script

This Short Python script takes advantage of the SQLalchemy and XlsxWriter modules to connect to multiple database flavours, fetch some data according to a SQL Query and populate an Excel sheet with the data

By design, it only takes one SELECT statement per script run, if data from multiple tables is needed, a UNION statement may be used.

Supported databases:
1. Sqlite
2. MySQL
3. PostgreSQL
4. Oracle
5. Microsoft SQL Server

<p align="right">(<a href="#top">back to top</a>)</p>


<!-- GETTING STARTED -->
## Getting Started

Clone the github repo or download the files manually

### Prerequisites

1. Create a virtual environment using virtualenv
  ```sh
  pip install virtualenv
  virtualenv db_to_xlsx_venv
  ```

2. Activate virtual environment
  ```sh
  source db_to_xlsx_venv/bin/activate
  ```
3. Install the requirements from requirements.txt
  ```sh
  pip install -r requirements.txt
  ```

<!-- USAGE EXAMPLES -->
## Usage
  ```sh
  python3 db_to_xlsx.py [-h] -f {sqlite,mysql,postgresql,oracle,mssql} [-u USERNAME] [-pw PASSWORD] [-hn HOSTNAME] [-p PORT] -db DATABASE 
  ```
By design choice, the input parameters are given as command line arguments while the sql statement is given in INPUT form so as to establish connection first. Optional parameters are *user*, *pass*, *host*, *port*. Non-Optional parameters are *flavour*, *database*. Usage examples for a couple of database flavours:

1. SQLite
  ```sh
  python3 db_to_xlsx.py -f sqlite -db example1.db
  ```
  
2. MySQL
  ```sh
  python3 db_to_xlsx.py -f mysql -u user -pw pass -hn localhost -p 3306 -db example2.db
  ```

Excel files are by default stored in "excel_files" directory

For more information type the following command for the help page
  ```sh
  python3 db_to_xlsx.py -h 
  ```

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- ROADMAP -->
## Testing progress

- [x] Tested Sqlite
  - [x] local
- [x] Tested MySQL
  - [x] local
  - [x] remote 
- [x] Tested PostgreSQL
  - [x] local
  - [x] remote 
- [] Tested Oracle
  - [] local
  - [] remote 
- [] Tested MsSQL Server
  - [] local
  - [] remote


<p align="right">(<a href="#top">back to top</a>)</p>



