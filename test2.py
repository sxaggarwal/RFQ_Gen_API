import os
from src.general_class import TableManger
from src.schema import _get_schema

quote_assembly_table = TableManger("QuoteAssembly")
quote_assembly_formula_variable_table = TableManger(
            "QuoteAssemblyFormulaVariable"
        )

# all_columns = _get_schema("QuoteAssembly")
# ret_list = []
# columns_to_omit = [
#     "LastAccess",
# ]
# for row in all_columns:
#     if row[0] not in columns_to_omit:
#         ret_list.append(row[0])
# ret_list_str = ",".join(map(str, ret_list))
# temp = quote_assembly_table.get(ret_list_str, QuoteFK=1230)


# for data in temp:
#     info_dict = dict(zip(ret_list, data))

#     print(info_dict)
#     print("#" * 50)
#     quote_assembly_table.insert(info_dict, live=True)

temp = quote_assembly_table.get(
            "QuoteAssemblyPK",
            "SetupFormulaFK",
            "RunFormulaFK",
            "OperationFK",
            "SetupTime",
            "RunTime",
            QuoteFK=1230,
        )
for a, b, c, d, e, f in temp:
    if d:
        dict_1 = {
            "QuoteAssemblyFK": a,
            "OperationFormulaVariableFK": b,
            "FormulaType": 0,
            "VariableValue": e,  # can change this in future
        }
        dict_2 = {
            "QuoteAssemblyFK": a,
            "OperationFormulaVariableFK": c,
            "FormulaType": 1,
            "VariableValue": f,  # can change this in future
        }
        quote_assembly_formula_variable_table.insert(dict_1, live=True)
        quote_assembly_formula_variable_table.insert(dict_2, live=True)