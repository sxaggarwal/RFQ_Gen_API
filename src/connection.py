import pyodbc
from src.logging_config import setup_logging

_conn_string_sandbox = "DRIVER={SQL Server};SERVER=ETZ-SQL;DATABASE=SANDBOX;Trusted_Connection=yes"
_conn_string_live = "DRIVER={SQL Server};SERVER=ETZ-SQL;DATABASE=ETEZAZIMIETrakLive;Trusted_Connection=yes"

setup_logging()

def get_connection(live=False):
    """Default connection is sandbox, for live set live=True. Returns the connection object"""
    if live:
        return pyodbc.connect(_conn_string_live)
    return pyodbc.connect(_conn_string_sandbox)

# if __name__ == "__main__":
#     get_connection(live=True)