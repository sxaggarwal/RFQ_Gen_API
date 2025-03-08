from typing import Dict, Any

import pyodbc
from mie_trak_api.utils import with_db_conn
from base_logger import getlogger


LOGGER = getlogger("MT RFQ")


@with_db_conn(commit=True)
def insert_into_rfq(
    cursor: pyodbc.Cursor,
    customer_fk: int,
    address_dict: Dict[str, Any],
    customer_rfq_number=None,
    buyer_fk=None,
    inquiry_date=None,
    due_date=None,
    create_date=None,
    rfq_status_fk: int = 5,
) -> int:
    """
    [TODO:description]

    :param cursor: [TODO:description]
    :param customer_fk: [TODO:description]
    :param address_dict: [TODO:description]
    :param customer_rfq_number [TODO:type]: [TODO:description]
    :param buyer_fk [TODO:type]: [TODO:description]
    :param inquiry_date [TODO:type]: [TODO:description]
    :param due_date [TODO:type]: [TODO:description]
    :param create_date [TODO:type]: [TODO:description]
    :param rfq_status_fk: [TODO:description]
    :return: [TODO:description]
    :raises ValueError: [TODO:description]
    """
    info_dict = {
        "CustomerFK": customer_fk,
        "BuyerFK": buyer_fk,
        "BillingAddressFK": address_dict.get("address_pk"),
        "ShippingAddressFK": address_dict.get("address_pk"),
        "DivisionFK": 1,
        "ReceivedPurchaseOrder": 0,
        "NoBid": 0,
        "DidNotGet": 0,
        "MIEExchange": 0,
        "SalesTaxOnFreight": 0,
        "RequestForQuoteStatusFK": rfq_status_fk,
        "BillingAddressName": address_dict.get("address1"),
        "BillingAddress1": address_dict.get("address1"),
        "BillingAddress2": address_dict.get("address2"),
        "BillingAddressAlt": address_dict.get("address_alt"),
        "BillingAddressCity": address_dict.get("city"),
        "BillingAddressZipCode": address_dict.get("zip_code"),
        "ShippingAddressName": address_dict.get("address1"),
        "ShippingAddress1": address_dict.get("address1"),
        "ShippingAddress2": address_dict.get("address2"),
        "ShippingAddressAlt": address_dict.get("address_alt"),
        "ShippingAddressCity": address_dict.get("city"),
        "ShippingAddressZipCode": address_dict.get("zip_code"),
        "BillingAddressStateDescription": address_dict.get("state"),
        "BillingAddressCountryDescription": address_dict.get("country"),
        "ShippingAddressStateDescription": address_dict.get("state"),
        "ShippingAddressCountryDescription": address_dict.get("country"),
        "CustomerRequestForQuoteNumber": customer_rfq_number,
        "InquiryDate": inquiry_date,
        "DueDate": due_date,
        "CreateDate": create_date,
    }

    # Remove None values to prevent SQL errors
    filtered_dict = {k: v for k, v in info_dict.items() if v is not None}

    columns = ", ".join(filtered_dict.keys())
    placeholders = ", ".join(["?"] * len(filtered_dict))
    values = tuple(filtered_dict.values())

    query = f"""
    INSERT INTO RequestForQuote ({columns})
    VALUES ({placeholders})
    """

    cursor.execute(query, values)
    cursor.execute("SELECT IDENT_CURRENT('RequestForQuote')")
    result = cursor.fetchone()

    if not result or not result[0]:
        raise ValueError("RFQ PK was not returned by the database.")

    return int(result[0])


@with_db_conn(commit=True)
def reset_rfq(cursor: pyodbc.Cursor, rfq_pk: int) -> None:
    """
    Deletes all RFQ line items, associated quotes, and quote assemblies for a given RFQ.

    :param rfq_pk: The RFQ primary key.
    """
    query = """
        DELETE FROM QuoteAssembly
        WHERE QuoteFK IN (
            SELECT QuoteFK FROM RequestForQuoteLine WHERE RequestForQuoteFK = ?
        );

        DELETE FROM Quote
        WHERE QuotePK IN (
            SELECT QuoteFK FROM RequestForQuoteLine WHERE RequestForQuoteFK = ?
        );

        DELETE FROM RequestForQuoteLine
        WHERE RequestForQuoteFK = ?;
    """

    cursor.execute(query, (rfq_pk, rfq_pk, rfq_pk))
    LOGGER.info("RFQ PK: {rfq_pk} reset successful.")


@with_db_conn(commit=True)
def create_rfq_line_item(
    cursor: pyodbc.Cursor,
    item_fk: int,
    request_for_quote_fk: int,
    line_reference_number: int,
    quote_fk: int,
    price_type_fk=3,
    unit_of_measure_set_fk=1,
    quantity=None,
):
    """
    Adds a line item to the RFQ using a single SQL query.

    :param item_fk: Foreign key reference to the Item table.
    :param request_for_quote_fk: Foreign key reference to the RFQ table.
    :param line_reference_number: The reference number for the RFQ line item.
    :param quote_fk: Foreign key reference to the Quote table.
    :param price_type_fk: The price type (defaults to 3).
    :param unit_of_measure_set_fk: Unit of measure set (defaults to 1).
    :param quantity: The requested quantity.
    :return: The primary key (PK) of the newly inserted RFQ line item.
    """

    query = """
        INSERT INTO RequestForQuoteLine
        (ItemFK, RequestForQuoteFK, LineReferenceNumber, QuoteFK, Quantity, PriceTypeFK, UnitOfMeasureSetFK)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """
    cursor.execute(
        query,
        (
            item_fk,
            request_for_quote_fk,
            line_reference_number,
            quote_fk,
            quantity,
            price_type_fk,
            unit_of_measure_set_fk,
        ),
    )
    cursor.execute("SELECT IDENT_CURRENT('RequestForQuoteLine')")
    result = cursor.fetchone()

    if not result or not result[0]:
        raise ValueError("RFQ PK was not returned by the database.")

    return int(result[0])


@with_db_conn(commit=True)
def create_rfq_line_item_with_qty(
    cursor: pyodbc.Cursor,
    item_fk: int,
    request_for_quote_fk: int,
    line_reference_number: int,
    quote_fk: int,
    quantity: float = 1.00,
    price_type_fk: int = 3,
    unit_of_measure_set_fk: int = 1,
    delivery: int = 1,
):
    """
    Inserts a new RFQ line item and its quantity in a single SQL transaction.

    :param item_fk: Foreign key reference to the Item table.
    :param request_for_quote_fk: Foreign key reference to the RFQ table.
    :param line_reference_number: The reference number for the RFQ line item.
    :param quote_fk: Foreign key reference to the Quote table.
    :param quantity: The requested quantity (defaults to 1.00).
    :param price_type_fk: The price type (defaults to 3).
    :param unit_of_measure_set_fk: Unit of measure set (defaults to 1).
    :param delivery: Delivery ID (defaults to 1).
    :return: The primary key (PK) of the newly inserted RFQ line item.
    """

    # Step 1: Insert into RequestForQuoteLine and retrieve the inserted PK
    insert_rfq_line_query = """
        INSERT INTO RequestForQuoteLine 
        (ItemFK, RequestForQuoteFK, LineReferenceNumber, QuoteFK, Quantity, PriceTypeFK, UnitOfMeasureSetFK)
        VALUES (?, ?, ?, ?, ?, ?, ?);
    """

    cursor.execute(
        insert_rfq_line_query,
        (
            item_fk,
            request_for_quote_fk,
            line_reference_number,
            quote_fk,
            quantity,
            price_type_fk,
            unit_of_measure_set_fk,
        ),
    )

    cursor.execute("SELECT IDENT_CURRENT('RequestForQuoteLine')")
    result = cursor.fetchone()

    if not result or not result[0]:
        raise ValueError("RFQ Line PK was not returned by the database.")

    rfq_line_pk = result[0]
    LOGGER.debug(f"Inserted RFQ Line PK: {rfq_line_pk}")

    # Step 2: Insert into RequestForQuoteLineQuantity using the retrieved PK
    insert_rfq_line_qty_query = """
        INSERT INTO RequestForQuoteLineQuantity 
        (RequestForQuoteLineFK, PriceTypeFK, Quantity, Delivery)
        VALUES (?, ?, ?, ?);
    """

    cursor.execute(
        insert_rfq_line_qty_query,
        (rfq_line_pk, price_type_fk, quantity, delivery),
    )
    LOGGER.debug("Qty Updated.")

    return rfq_line_pk
