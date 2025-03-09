import re
from typing import Dict, Any
from base_logger import getlogger
from mie_trak_api import request_for_quote, quote, item, router


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


def create_finish_router(finish_description: str, item_fin_pk: int, part_num: str):
    "Adds a router for every finish"
    finish_code = finish_description.split("\n")
    finish_pks = []

    if finish_code:
        for code in finish_code:
            finish_codes_pk = item.get_or_create_item(
                **{
                    "PartNumber": code[:100],
                    "Description": code[
                        :490
                    ],  # TODO: Fix this as its crossing the limit, add this to the comments.
                    "Inventoriable": 0,
                    "ItemTypeFK": 5,
                    "CertReqdBySupplier": 1,
                    "CanNotCreateWorkOrder": 1,
                    "CanNotInvoice": 1,
                    "PurchaseAccountFK": 125,
                    "CogsAccFk": 125,
                    "CalculationTypeFK": 17,
                    "Comment": code,
                }
            )
            finish_pks.append(finish_codes_pk)

    router_pk = router.create_router(item_fin_pk, part_num)
    LOGGER.debug(f"Created Router PK: {router_pk}")

    for idx, pk in enumerate(finish_pks, start=1):
        router.create_router_work_center(pk, router_pk, idx)


def center_window(window, width=1000, height=700):
    """
    [TODO:description]

    :param window [TODO:type]: [TODO:description]
    :param width [TODO:type]: [TODO:description]
    :param height [TODO:type]: [TODO:description]
    """
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)

    window.geometry(f"{width}x{height}+{x}+{y}")
