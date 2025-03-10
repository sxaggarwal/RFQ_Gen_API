from src.connection import get_connection
from src.exceptions import  SchemaError
from src.schema import _get_schema, print_schema
import pyodbc
import logging


class TableManger:
    """TableManager that processes all table requests
    Pass table name in the constructor. Use insert, get, delete to get data.
    Import print_schema from schema and use self.schema to print with better visibility.
    Add any mandetory field to insert_mandetory or insert_not_allowed lists to run checks on them.
    """

    def __init__(self, table_name):
        """Table_name: name of the table in Mie Trak database"""
        assert table_name not in ["item"]
        self.table_name = table_name
        self.logger = logging.getLogger().getChild(f"GeneralAPI - {self.table_name}")
        self.schema = _get_schema(self.table_name)
        self.column_names = [name[0] for name in self.schema]
        self.insert_not_allowed = [f"{self.table_name}PK", ]
        self.insert_mandetory = []

        self.logger.info(f"init new. Table Name - {self.table_name}")

    def _column_check(self, columns, insert=False):
        """Checks if the columns being accessed exist in the schema. Insert is False by default"""
        for value in columns:
            if value not in self.column_names:
                raise SchemaError.column_does_not_exist_error(value)
            elif insert:
                if value in self.insert_not_allowed:
                    raise SchemaError.insertion_not_allowed_error(value)
        if insert:
            for val in self.insert_mandetory:
                if val not in columns:
                    raise SchemaError.mandetory_column_missing_error(val, self.table_name)   

    def insert(self, update_dict: dict):
        """update_dict: key-value pairs of column_names: values. These values will be checked against the schema before they go in the table"""
        self._column_check(update_dict.keys(), insert=True)
        column_names, column_len = ", ".join([val for val in update_dict.keys()]), ", ".join(["?" for val in update_dict.keys()])
        values = [val for val in update_dict.values()]
        query = f"""INSERT INTO {self.table_name} ({column_names})
                    VALUES ({column_len})"""
        try:
            with get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, *values)
                query = f"SELECT IDENT_CURRENT('{self.table_name}')"
                cursor.execute(query)
                return_pk = cursor.fetchone()[0]
                self.logger.info(f"Inserted {self.table_name}. {self.table_name}PK: {return_pk}")
                conn.commit()
            return return_pk
        except pyodbc.Error as e:
            self.logger.error(f"Inserted {self.table_name}. {self.table_name}PK: {return_pk} - {e}")
            print(e)

    def get(self, *args, **kwargs) -> list:
        """args: define what is returned as a tuple
            kwargs: define what is passed as a parameter"""
        self._column_check(kwargs.keys())
        return_param_string = ",".join(args) if args else "*" 
        query = f"SELECT {return_param_string} FROM {self.table_name}"
        parameters = []
        if kwargs:
            search_params = []
            for key, value in kwargs.items():
                if value is None:
                    search_params.append(f"{key} IS NULL")
                else:
                    search_params.append(f"{key} = ?")
                    parameters.append(value)
            query += " WHERE " + " AND ".join(search_params)
        self.logger.info(f"GET METHOD Query Built -> {query}")
        try:
            with get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, parameters)
                result = cursor.fetchall()
                if result:
                    return result
                else:
                    return None
                    # raise ItemNotFoundError        
        except pyodbc.Error as e:
            self.logger.error(f"GET METHOD Query Built -> {query} -> ERROR {e}")
            print(e)

    def update(self, pk, **kwargs):
        """pk to get the item to update, kwargs to update the columns"""
        if not kwargs:
            print("Nothing updated, no kwargs passed")
            return  # Exit the function if no kwargs are provided

        set_string = ", ".join([f"{key} = '{value}'" for key, value in kwargs.items()])

        query = f"UPDATE {self.table_name} SET {set_string} WHERE {self.table_name}PK = {pk};"
        self.logger.info(f"UPDATE METHOD Query Built -> {query}")
        try:
            with get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
        except pyodbc.Error as e:
            self.logger.error(f"GET METHOD Query Built -> {query}")
            print(f"Error - pyodbc -> {e}")

    def delete(self, pk):
        """Takes PK and deletes that entry"""
        try:
            with get_connection() as conn:
                cursor = conn.cursor()
                query = f"DELETE FROM {self.table_name} WHERE {self.table_name}PK=?;"
                self.logger.info(f"DELETE - {self.table_name}PK: {pk}")
                cursor.execute(query, pk)
                conn.commit()
        except pyodbc.Error as e:
            print(e)

if __name__ == "__main__":
    i_table = TableManger("Item")
    print_schema(i_table.schema)