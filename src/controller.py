import re
from typing import Dict, Any
from base_logger import getlogger
from mie_trak_api import request_for_quote, quote


LOGGER = getlogger("Controller")


def create_rfq(
    quote_pk_dict,
    item_pk_dict,
    rfq_pk,
    info_dict: Dict[str, Dict[str, Any]],
    parent_quote_fk=None,
    i=1,
):
    """checks if its Assy or Detail and accordingly creates the line item and adds quotes of assembly to the BOM of Assy Line Quotes"""
    main_part_number = None
    main_quote_pk = None

    parent_quote_assembly_pk_dict = {}
    for new_key, value in info_dict.items():
        if re.search(r"_____\d+$", new_key) is None:
            key = new_key
        else:
            key = new_key.split("_____")[0]

        part_number = key
        assy_for = value.get("assy_for", None)
        quote_pk = quote_pk_dict.get(part_number)
        item_pk = item_pk_dict.get(part_number)

        if not assy_for:
            rfq_line_pk = request_for_quote.create_rfq_line_item_with_qty(
                item_pk,
                rfq_pk,
                i,
                quote_pk,
                quantity=value.get("quantity_required"),
            )
            LOGGER.debug(f"{rfq_line_pk}")
            i += 1
            main_quote_pk = quote_pk
            main_part_number = part_number

        elif assy_for and not value.get("hardware_or_supplies"):
            LOGGER.debug("executing hardware or supplies...")

            if not main_part_number or not main_quote_pk:
                raise ValueError("Data from excel sheet is not proper bruh.")

            if assy_for == main_part_number:
                quote_fk = main_quote_pk
                parent_quote_assembly_pk = quote.create_assy_quote(
                    quote_pk, quote_fk, value.get("quantity_required", "")
                )
                parent_quote_assembly_pk_dict[part_number] = parent_quote_assembly_pk

            else:
                LOGGER.debug("executing else in create RFQ.")
                parent_quote_fk = quote_pk_dict[assy_for]
                if assy_for not in parent_quote_assembly_pk_dict:
                    raise KeyError(
                        f"Key '{assy_for}' not found in parent_quote_assembly_pk_dict"
                    )
                parent_quote_assembly_pk_new = parent_quote_assembly_pk_dict[assy_for]
                parent_quote_assembly_pk = quote.create_assy_quote(
                    quote_pk,
                    main_quote_pk,
                    value.get("quantity_required", 1),
                    parent_quote_fk=parent_quote_fk,
                    parent_quote_asembly=parent_quote_assembly_pk_new,
                )
                parent_quote_assembly_pk_dict[part_number] = parent_quote_assembly_pk
