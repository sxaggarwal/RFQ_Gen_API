from src import excel_parser
from pprint import pprint
from mie_trak_api import quote, item
from src.helper import pk_info_dict, create_dict_from_excel
from src.excel_parser import create_dict_from_excel_new


if __name__ == "__main__":
    filepath = r"C:\Users\svyas\Downloads\CA-1328893.xlsx"
    info_dict = create_dict_from_excel(filepath)

    d2 = create_dict_from_excel_new(filepath)
    pprint(d2)

    pprint(info_dict)
