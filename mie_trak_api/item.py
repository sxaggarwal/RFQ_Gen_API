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

    validated_data = item_model(**item_data).model_dump(exclude_unset=True)

    part_number = validated_data.get("PartNumber")
    cursor.execute("SELECT ItemPK FROM Item WHERE PartNumber = ?", (part_number, ))
    result = cursor.fetchone()

    if result:
        LOGGER.info(f"PartNumber: {part_number} found. (PK: {result[0]})")
        return result[0]

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
    LOGGER.info(f"Query created:\n{query}")

    cursor.execute(query, values)
    cursor.execute("SELECT SCOPE_IDENTITY()")
    result = cursor.fetchone()

    if result:
        return result[0]
    else:
        raise ValueError("`SELECT SCOPE` did not return anything. Item might not be inserted.")

