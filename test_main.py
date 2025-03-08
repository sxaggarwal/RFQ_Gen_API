from src import excel_parser
from pprint import pprint
from mie_trak_api import quote, item, request_for_quote
from src.mie_trak import MieTrak


if __name__ == "__main__":
    # request_for_quote.reset_rfq(5265)
    column_names, template_values = quote.get_operation_quote_template()
    # m = MieTrak()
    # col_name, temp_vals = m.quote_operation_template()
    for data in template_values:
        insert_dict = dict(zip(column_names, data))
        print(insert_dict)
        insert_dict["QuoteFK"] = 1
        insert_dict["ParentQuoteAssemblyFK"] = 1
        insert_dict["ParentQuoteFK"] = 1
        insert_columns = ", ".join(list(insert_dict.keys()))
        placeholders = ", ".join(["?"] * len(insert_dict))
        insert_query = (
            f"INSERT INTO QuoteAssembly ({insert_columns}) VALUES ({placeholders})"
        )
        print(insert_query)
        break
        # print(insert_query)
