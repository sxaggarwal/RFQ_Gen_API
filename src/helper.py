from src.general_class import TableManger
from src.mie_trak import MieTrak
import os
import shutil
import math
import pandas as pd

def transfer_file_to_folder(folder_path: str, file_path: str) -> str:
        """ Copies file from one path to another path"""
        os.makedirs(folder_path, exist_ok=True)

        filename = os.path.basename(file_path)  # source file path
        destination_path = os.path.join(folder_path, filename)
        shutil.copyfile(file_path, destination_path)

        return destination_path

def extract_from_excel(filepath, column_name):
    """ Extracting a column from an excel file"""
    df = pd.read_excel(filepath)
    data = df[column_name].tolist()
    return data

def create_dict_from_excel(filepath):
    part_number = extract_from_excel(filepath, "Part")
    description = extract_from_excel(filepath, "DESCRIPTION")
    length = extract_from_excel(filepath, "Length")
    thickness = extract_from_excel(filepath, "Thickness")
    width = extract_from_excel(filepath, "Width")
    weight = extract_from_excel(filepath, "Weight")
    material = extract_from_excel(filepath, "Material")
    finish_code = extract_from_excel(filepath, "FinishCode")
    heat_treat = extract_from_excel(filepath, "HeatTreat")
    drawing_number = extract_from_excel(filepath, "DrawingNumber")
    drawing_revision = extract_from_excel(filepath, "DrawingRevision")
    qty_reqd = extract_from_excel(filepath, "QuantityRequired")
    pl_rev = extract_from_excel(filepath, "PLRevision")
    assy_for = extract_from_excel(filepath, "AssyFor")
    hardware_or_supplies = extract_from_excel(filepath, "Hardware/Supplies")
    my_dict = {}

    for a,b,c,d,e,f,g,h,i,j,k,l,m,n,o in zip(part_number,description,length,thickness,width,weight,material,finish_code,heat_treat,drawing_number,drawing_revision,qty_reqd,pl_rev,assy_for, hardware_or_supplies):  # noqa: E741
        a = None if isinstance(a, float) and math.isnan(a) else a 
        b = None if isinstance(b, float) and math.isnan(b) else b
        c = 0.00 if isinstance(c, float) and math.isnan(c) else c
        d = 0.00 if isinstance(d, float) and math.isnan(d) else d
        e = 0.00 if isinstance(e, float) and math.isnan(e) else e
        f = 0.00 if isinstance(f, float) and math.isnan(f) else f
        g = None if isinstance(g, float) and math.isnan(g) else g
        h = None if isinstance(h, float) and math.isnan(h) else h
        i = None if isinstance(i, float) and math.isnan(i) else i
        j = None if isinstance(j, float) and math.isnan(j) else j
        k = None if isinstance(k, float) and math.isnan(k) else k
        l = None if isinstance(l, float) and math.isnan(l) else l  # noqa: E741
        m = None if isinstance(m, float) and math.isnan(m) else m
        n = None if isinstance(n, float) and math.isnan(n) else n
        o = None if isinstance(o, float) and math.isnan(o) else o

        my_dict[a] = (b,c,d,e,f,g,h,i,j,k,l,m,n,o)
    
    return my_dict

def pk_info_dict(info_dict):
    item_table = TableManger("Item")
    m = MieTrak()
    my_dict = {}
    for key, value in info_dict.items():
        mat_pk = None
        fin_pk = None
        ht_pk = None
        if value[5]:
            result = item_table.get("ItemPK", PartNumber=value[5], StockLength=value[1], StockWidth=value[3])
            if result:
                mat_pk = result[0][0]
            else:
                pk = m.get_or_create_item(value[5], service_item=0, purchase=1, manufactured_item=0, item_type_fk= 2, only_create = 1, bulk_ship=0, ship_loose=0, cert_reqd_by_supplier=1, purchase_account_fk=127, cogs_acc_fk=127, calculation_type_fk=4)
                mat_pk = pk
        
        if value[6]:
            result = item_table.get("ItemPK", PartNumber=f"{key} - OP Finish")
            if result: 
                fin_pk = result[0][0]
            else:
                pk = m.get_or_create_item(part_number=f"{key} - OP Finish", item_type_fk=5, comment = value[6], purchase_order_comment=value[6], inventoriable=0, only_create=1, cert_reqd_by_supplier=1, can_not_create_work_order=1, can_not_invoice=1, purchase_account_fk=125, cogs_acc_fk=125, calculation_type_fk=17)
                fin_pk = pk
        
        if value[7]:
            result = item_table.get("ItemPK", PartNumber=f"{key} - OP HT")
            if result: 
                ht_pk = result[0][0]
            else:
                pk = m.get_or_create_item(part_number=f"{key} - OP HT", item_type_fk=5, description= value[7], comment= value[7], purchase_order_comment= value[7], inventoriable= 0, only_create=1, cert_reqd_by_supplier=1, can_not_create_work_order=1, can_not_invoice=1, purchase_account_fk=125, cogs_acc_fk=125, calculation_type_fk=17)
                ht_pk = pk
        my_dict[key] = (mat_pk, ht_pk, fin_pk)
    return my_dict
          
