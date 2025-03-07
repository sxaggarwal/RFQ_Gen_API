import pyodbc
from mie_trak_api.utils import get_table_schema, with_db_conn
from base_logger import getlogger


LOGGER = getlogger("MT Quote")


@with_db_conn(commit=True)
def create_quote_new(cursor: pyodbc.Cursor, customer_fk: int, item_fk: int, quote_type: int, part_number: str):
    """
    [TODO:description]

    :param cursor: [TODO:description]
    :param customer_fk: [TODO:description]
    :param item_fk: [TODO:description]
    :param quote_type: [TODO:description]
    :param part_number: [TODO:description]
    :raises ValueError: [TODO:description]
    """
    query = """
        INSERT INTO Quote (CustomerFK, ItemFK, QuoteType, PartNumber, DivisionFK) 
        VALUES (?, ?, ?, ?, ?);
    """
    cursor.execute(query, (customer_fk, item_fk, quote_type, part_number, 1))
    cursor.execute("SELECT IDENT_CURRENT('Quote');")
    result = cursor.fetchone()

    if not result or result[0] is None:
        raise ValueError("Quote PK was not returned by the database.")

    return int(result[0])


@with_db_conn(commit=True)
def copy_operations_to_quote(cursor: pyodbc.Cursor, new_quote_fk, source_quote_fk=494):
    """
    [TODO:description]

    :param cursor: [TODO:description]
    :param new_quote_fk [TODO:type]: [TODO:description]
    :param source_quote_fk [TODO:type]: [TODO:description]
    """

    all_columns = get_table_schema("QuoteAssembly")
    excluded_columns = ["QuoteFK", "QuoteAssemblyPK", "LastAccess", "ParentQuoteAssemblyFK", "ParentQuoteFK"]

    columns_to_copy = [column.get("column_name") for column in all_columns if column.get("column_name") not in excluded_columns]
    column_names = ", ".join(columns_to_copy)  # Convert list to SQL-friendly format

    query = f"""
        INSERT INTO QuoteAssembly ({column_names}, QuoteFK)
        SELECT {column_names}, ?
        FROM QuoteAssembly
        WHERE QuoteFK = ?;
    """

    cursor.execute(query, (new_quote_fk, source_quote_fk))
    LOGGER.info(f"Copied QuotePK: {source_quote_fk} to NEW QuotePK: {new_quote_fk}")


@with_db_conn()
def get_quote_assembly_pk(cursor: pyodbc.Cursor, **quote_details) -> int | None:
    """
    [TODO:description]

    :param cursor: [TODO:description]
    :return: [TODO:description]
    :raises ValueError: [TODO:description]
    """
    if not quote_details:
        raise ValueError("At least one condition must be provided to get an item.")

    where_conditions = " AND ".join([f"{key} = ?" for key in quote_details.keys()])
    query = f"SELECT QuoteAssemblyPK FROM QuoteAssembly WHERE {where_conditions};"

    values = tuple(quote_details.values())

    cursor.execute(query, values)
    result = cursor.fetchone()

    return result[0] if result else None
