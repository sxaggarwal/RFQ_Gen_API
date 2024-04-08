from pyodbc import Error

# we need this as pyodbc will happily return an empty fetchall list
class ItemNotFoundError(Exception):
    def __init__(self, part_number, message="Search value does not exist in item table, cannot update."):
        self.part_number = part_number
        self.message = message
        super().__init__(f"{self.message} Column Name: {self.part_number}")

class SchemaError(Exception):
    def __init__(self, column_name, message="Schema Error"):
        self.column_name = column_name
        self.message = message
        super().__init__(f"{self.message} Column Name: {self.column_name}")

    @staticmethod
    def insertion_not_allowed_error(column_name):
        raise SchemaError(column_name, message="Insert into column not allowed.") 

    @staticmethod
    def column_does_not_exist_error(column_name):
        raise SchemaError(column_name, message="This column name does not exist")

    @staticmethod
    def mandetory_column_missing_error(column_name, table_name):
        raise SchemaError(column_name, message=f"{column_name} is missing. This is mandetory for {table_name}")

class TableDoesNotExistError(Exception):
    def __init__(self, table_name, message="Table does not exist"):
        self.table_name = table_name
        self.message = message
        super().__init__(f"{self.table_name} {self.message}")

if __name__ == "__main__":
    raise TableDoesNotExistError("1234") 
