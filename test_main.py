from src import excel_parser
from pprint import pprint
from mie_trak_api import quote, item
from src.helper import pk_info_dict, create_dict_from_excel


if __name__ == "__main__":
    # filepath = r"C:\Users\svyas\Downloads\CA-1328893.xlsx"
    # info_dict = create_dict_from_excel(filepath)
    # info_dict_new = excel_parser.create_dict_from_excel_new(filepath)
    # # pprint(info_dict)
    # pprint(info_dict_new)
    #
    # pprint(pk_info_dict(info_dict_new))

    # pprint(pk_info_dict(info_dict))

    # print(item.get_item(**{"PartNumber": "B02-20697-10 - OP Finish"}))
    print(type(quote.get_quote_assembly_pk(**{"QuoteFK": 2539, "SequenceNumber": 21})))

