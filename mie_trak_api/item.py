from typing import ValuesView
from mie_trak_api.utils import with_db_conn, create_pydantic_model
from base_logger import getlogger
import pyodbc


LOGGER = getlogger("MT Item")

item_model = create_pydantic_model("item")


@with_db_conn(commit=True)
def get_or_create_item(cursor: pyodbc.Cursor, **item_data):
    """
    [TODO:description]

    :param cursor: [TODO:description]
    :raises ValueError: [TODO:description]
    :raises ValueError: [TODO:description]
    :raises ValueError: [TODO:description]
    """

    if not item_data.get("PartNumber", None):
        raise ValueError("kwargs must contain a part number")

    part_number = item_data.get("PartNumber")
    cursor.execute("SELECT ItemPK FROM Item WHERE PartNumber = ?", (part_number,))
    result = cursor.fetchone()

    if result:
        LOGGER.info(f"PartNumber: {part_number} found. (PK: {result[0]})")
        return result[0]

    validated_data = item_model(**item_data).model_dump(exclude_unset=True)

    cursor.execute("INSERT INTO ItemInventory (QuantityOnHand) Values (0.000)")
    cursor.execute("SELECT SCOPE_IDENTITY()")
    item_inventory_pk = cursor.fetchone()

    if not item_inventory_pk:
        raise ValueError("ItemPK was not returned.")

    validated_data["ItemInventoryFK"] = int(item_inventory_pk[0])

    columns = ", ".join(validated_data.keys())
    placeholders = ", ".join(["?"] * len(validated_data))
    values = tuple(validated_data.values())

    query = f"INSERT INTO Item ({columns}) VALUES ({placeholders})"

    cursor.execute(query, values)
    cursor.execute("SELECT IDENT_CURRENT('Item')")
    result = cursor.fetchone()

    if result and result[0]:
        LOGGER.info(f"Inserted new ItemPK: {result[0]}")
        return result[0]
    else:
        LOGGER.critical("SELECT IDENT failed in get_or_create_item")
        raise ValueError(
            "`SELECT SCOPE` did not return anything. Item might not be inserted."
        )


@with_db_conn()
def get_item(cursor: pyodbc.Cursor, **item_data) -> int | None:
    """
    [TODO:description]

    :param cursor: [TODO:description]
    :return: [TODO:description]
    :raises ValueError: [TODO:description]
    """
    if not item_data:
        raise ValueError("At least one condition must be provided to get an item.")

    where_conditions = " AND ".join([f"{key} = ?" for key in item_data.keys()])
    query = f"SELECT ItemPK FROM Item WHERE {where_conditions};"

    values = tuple(item_data.values())

    cursor.execute(query, values)
    result = cursor.fetchone()

    return result[0] if result else None


@with_db_conn(commit=True)
def update_item(cursor, itempk: int, **item_data) -> None:
    if not item_data:
        raise ValueError("At least one condition must be provided to get an item.")

    set_string = ", ".join([f"{key} = '{value}'" for key, value in item_data.items()])
    query = f"UPDATE Item SET {set_string} WHERE ItemPK = {itempk};"

    LOGGER.debug(query)
    cursor.execute(query)
    LOGGER.info(f"Updated ItemPK: {itempk}.")
