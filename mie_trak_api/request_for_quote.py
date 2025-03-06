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
    rfq_status_fk: int=5,
) -> int:
    """Inserts all the details and returns a RFQ number"""
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

