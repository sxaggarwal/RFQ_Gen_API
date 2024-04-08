from src.connection import get_connection
from src.exceptions import TableDoesNotExistError


def _get_schema(table_name: str):
    """Provides all the schema for given table name. Includes all columns"""
    with get_connection() as conn:
        cursor = conn.cursor()
        query = """
        SELECT 
            COLUMN_NAME, 
            DATA_TYPE, 
            CHARACTER_MAXIMUM_LENGTH, 
            IS_NULLABLE, 
            COLUMN_DEFAULT
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?;
        """
        cursor.execute(query, table_name)
        schema = cursor.fetchall()
        if schema:
            return schema
        else:
            raise TableDoesNotExistError(table_name)


def print_schema(schema):
    """Assumes that we will only be printing the schema from the _get_schema function"""
    print(f'{"Column Name":40} {"Data Type":20} {"Max Length":10} {"Nullable":10} {"Default":10}')
    for column_details in schema:
        print(f'{column_details[0]:40} {column_details[1]:20} {str(column_details[2]):10} {column_details[3]:10} {str(column_details[4]):10}')        

# WIP
# def prettify(list_of_tuples):
#     for x in range(len(list_of_tuples)):
#         print()
