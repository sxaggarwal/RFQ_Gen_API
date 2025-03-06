from src import excel_parser
from pprint import pprint
from mie_trak_api.item import get_or_create_item


if __name__ == "__main__":
    # filepath = r"C:\Users\svyas\Downloads\CA-1328893.xlsx"
    # filepath_error = r"C:\Users\svyas\Downloads\CA-1328893.xlsx"
    #
    # d2 = excel_parser.create_dict_from_excel_new(filepath_error)
    # for key, value in d2.items():
    #     data = {
    #         "PartNumber": key,
    #         "Description":value.get('description', ""),
    #     }
    #     get_or_create_item(**data)
    #     break


    data = {
        "PartNumber": "01-10422",
    }
    get_or_create_item(**data)
