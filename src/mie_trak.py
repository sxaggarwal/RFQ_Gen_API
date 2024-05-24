import os
from src.general_class import TableManger
from src.schema import _get_schema

class MieTrak:
    def __init__(self):
        self.party_table = TableManger("Party")
        self.address_table = TableManger("Address")
        self.state_table = TableManger("State")
        self.country_table = TableManger("Country")
        self.request_for_quote_table = TableManger("RequestForQuote")
        self.item_table = TableManger("Item")
        self.item_inventory_table = TableManger("ItemInventory")
        self.document_table = TableManger("Document")
        self.quote_table = TableManger("Quote")
        self.quote_assembly_table = TableManger("QuoteAssembly")
        self.rfq_line_table = TableManger("RequestForQuoteLine")
        self.rfq_line_qty_table = TableManger("RequestForQuoteLineQuantity")
        self.party_buyer_table = TableManger("PartyBuyer")
        self.router_table = TableManger("Router")
        self.quote_assembly_formula_variable_table = TableManger("QuoteAssemblyFormulaVariable")
        self.router_work_center_table = TableManger("RouterWorkCenter")

    def get_customer_data(self,selected_customer_index=None, names=False):
        data = self.party_table.get("PartyPK","Name", Customer=1)
        if names is True:
            customer_names = [d[1] for d in data]
            return customer_names
        else:
            customer_number_to_partypk = {i: d[0] for i,d in enumerate(data, start=1)}
            party_pk = customer_number_to_partypk[selected_customer_index+1]
            selected_customer_data = self.party_table.get("ShortName", "Email", PartyPK=party_pk)
            if selected_customer_data:
                short_name, email = selected_customer_data[0]
            else:
                short_name = None
                email = None
            return short_name, email, party_pk
    
    def get_buyer_info(self, buyer_fk):
        selected_customer_data = self.party_table.get("ShortName", "Email", PartyPK=buyer_fk)
        if selected_customer_data:
            short_name, email = selected_customer_data[0]
        else:
            short_name = None
            email = None
        return short_name, email
    
    def get_buyer_data(self, party_pk,):
        data = self.party_buyer_table.get("BuyerFK", PartyFK=party_pk)
        my_dict = {}
        if data:
            buyer_fk = [d[0] for d in data]
            if buyer_fk:
                for fk in buyer_fk:
                    buyer_name = self.party_table.get("Name", PartyPK=fk)[0][0]
                    my_dict[buyer_name] = fk
        return my_dict
    
    def get_address(self, party_fk):
        billing_details = self.address_table.get("AddressPK", "Name", "Address1", "Address2", "AddressAlt", "City", "ZipCode", PartyFK=party_fk)
        if billing_details:
            billing_details = billing_details[0]
        else:
            billing_details = [(None,None,None,None,None,None,None,)][0]
        info = self.address_table.get("StateFK", "CountryFK", PartyFK=party_fk)
        if info:
            state_fk, country_fk = info[0]
        else:
            state_fk=None
            country_fk=None
        if state_fk:
            state = self.state_table.get("Description", StatePK=state_fk)
        else:
            state = [(None,),] 

        if country_fk:
            country = self.country_table.get("Description", CountryPK=country_fk)
        else:
            country = [(None,),]   
        return billing_details, state, country
    
    def insert_into_rfq(self, customer_fk, billing_details, state, country, customer_rfq_number = None, buyer_fk = None, inquiry_date=None, due_date=None, create_date=None):
        info_dict = {
            "CustomerFK" : customer_fk,
            "BuyerFK" : buyer_fk,
            "BillingAddressFK": billing_details[0],
            "ShippingAddressFK": billing_details[0],
            "DivisionFK": 1,
            "ReceivedPurchaseOrder": 0,
            "NoBid": 0, 
            "DidNotGet": 0, 
            "MIEExchange": 0, 
            "SalesTaxOnFreight": 0,
            "RequestForQuoteStatusFK": 1,
            "BillingAddressName": billing_details[1], 
            "BillingAddress1": billing_details[2], 
            "BillingAddress2": billing_details[3], 
            "BillingAddressAlt": billing_details[4], 
            "BillingAddressCity": billing_details[5], 
            "BillingAddressZipCode": billing_details[6], 
            "ShippingAddressName": billing_details[1], 
            "ShippingAddress1": billing_details[2], 
            "ShippingAddress2": billing_details[3], 
            "ShippingAddressAlt": billing_details[4],
            "ShippingAddressCity": billing_details[5], 
            "ShippingAddressZipCode": billing_details[6], 
            "BillingAddressStateDescription": state[0][0],
            "BillingAddressCountryDescription": country[0][0], 
            "ShippingAddressStateDescription": state[0][0],
            "ShippingAddressCountryDescription": country[0][0], 
            "CustomerRequestForQuoteNumber": customer_rfq_number,
            "InquiryDate": inquiry_date,
            "DueDate": due_date,
            "CreateDate": create_date,
        }
        rfq_pk = self.request_for_quote_table.insert(info_dict)
        return rfq_pk
    
    def get_or_create_item(self, part_number, item_type_fk=1, mps_item=1, purchase=1, forecast_on_mrp=1, mps_on_mrp=1, service_item=1, unit_of_measure_set_fk=1, vendor_unit=1.0, manufactured_item=0, calculation_type_fk=1, inventoriable=1, purchase_order_comment=None,  description=None, comment=None, only_create=None, bulk_ship=1, ship_loose=1, cert_reqd_by_supplier=0, can_not_create_work_order=0, can_not_invoice=0, general_ledger_account_fk=100, purchase_account_fk=116, cogs_acc_fk=116):
        if only_create is None:
            item_pk = self.item_table.get("ItemPK", PartNumber=part_number)
            if item_pk:
                return item_pk[0][0]
            else:
                inventory_info_dict = {
                    "QuantityOnHand": 0.000,
                }
                item_inventory_pk = self.item_inventory_table.insert(inventory_info_dict)
                item_info_dict = {
                    "ItemInventoryFK" : item_inventory_pk, 
                    "PartNumber": part_number, 
                    "ItemTypeFK": item_type_fk, 
                    "Description" : description, 
                    "Comment": comment, 
                    "MPSItem": mps_item,
                    "Purchase": purchase, 
                    "ForecastOnMRP": forecast_on_mrp, 
                    "MPSOnMRP": mps_on_mrp, 
                    "ServiceItem": service_item, 
                    "PurchaseOrderComment": purchase_order_comment, 
                    "UnitOfMeasureSetFK": unit_of_measure_set_fk,
                    "VendorUnit": vendor_unit, 
                    "ManufacturedItem": manufactured_item, 
                    "CalculationTypeFK": calculation_type_fk, 
                    "Inventoriable": inventoriable, 
                    "BulkShip": bulk_ship, 
                    "ShipLoose": ship_loose,
                    "CertificationsRequiredBySupplier": cert_reqd_by_supplier,
                    "CanNotCreateWorkOrder": can_not_create_work_order,
                    "CanNotInvoice": can_not_invoice,
                    "GeneralLedgerAccountFK": general_ledger_account_fk,
                    "PurchaseGeneralLedgerAccountFK" : purchase_account_fk,
                    "SalesCogsAccountFK": cogs_acc_fk,
                }
                item_pk = self.item_table.insert(item_info_dict)
                return item_pk
        else:
            inventory_info_dict = {
                "QuantityOnHand": 0.000,
            }
            item_inventory_pk = self.item_inventory_table.insert(inventory_info_dict)
            item_info_dict = {
                "ItemInventoryFK" : item_inventory_pk, 
                "PartNumber": part_number, 
                "ItemTypeFK": item_type_fk, 
                "Description" : description, 
                "Comment": comment, 
                "MPSItem": mps_item,
                "Purchase": purchase, 
                "ForecastOnMRP": forecast_on_mrp, 
                "MPSOnMRP": mps_on_mrp, 
                "ServiceItem": service_item, 
                "PurchaseOrderComment": purchase_order_comment, 
                "UnitOfMeasureSetFK": unit_of_measure_set_fk,
                "VendorUnit": vendor_unit, 
                "ManufacturedItem": manufactured_item, 
                "CalculationTypeFK": calculation_type_fk, 
                "Inventoriable": inventoriable, 
                "BulkShip": bulk_ship, 
                "ShipLoose": ship_loose,
                "CertificationsRequiredBySupplier": cert_reqd_by_supplier,
                "CanNotCreateWorkOrder": can_not_create_work_order,
                "CanNotInvoice": can_not_invoice,
                "GeneralLedgerAccountFK": general_ledger_account_fk,
                "PurchaseGeneralLedgerAccountFK" : purchase_account_fk,
                "SalesCogsAccountFK": cogs_acc_fk,
            }
            item_pk = self.item_table.insert(item_info_dict)
            return item_pk
    
    def upload_documents(self, document_path: str, rfq_fk=None, item_fk=None, document_type_fk=None, secure_document=0, document_group_pk=None, print_with_purchase_order=None):
        if not rfq_fk and not item_fk:
            raise TypeError("Both values can not be None")
        found = False
        if item_fk:
            paths = self.document_table.get("URL", ItemFK=item_fk)
            if paths: 
                for path in paths:
                    if os.path.normpath(path[0]) == os.path.normpath(document_path):
                        found = True 
        if not found:
            doc_dict = {
                "URL": document_path,
                "RequestForQuoteFK": rfq_fk,
                "ItemFK": item_fk, 
                "Active": 1, 
                "DocumentTypeFK": document_type_fk, 
                "SecureDocument": secure_document, 
                "DocumentGroupFK": document_group_pk, 
                "PrintWithPurchaseOrder": print_with_purchase_order
            }
            self.document_table.insert(doc_dict)
    
    def create_quote(self, customer_fk, item_fk, quote_type, part_number):
        info_dict = {
            "CustomerFK": customer_fk, 
            "ItemFK": item_fk, 
            "QuoteType": quote_type, 
            "PartNumber": part_number, 
            "DivisionFK": 1,
        }
        quote_pk = self.quote_table.insert(info_dict)
        return quote_pk

    def quote_operation_template(self, quote_fk=494):
        all_columns = _get_schema("QuoteAssembly")
        ret_list = []
        columns_to_omit = ['QuoteFK', 'QuoteAssemblyPK', 'LastAccess', 'ParentQuoteAssemblyFK', 'ParentQuoteFK']
        for row in all_columns:
            if row[0] not in columns_to_omit:
                ret_list.append(row[0])
        
        if quote_fk == 494: 
            ret_list_str = ','.join(map(str, ret_list))
            temp = self.quote_assembly_table.get(ret_list_str, QuoteFK=494, ItemFK=None)
        else:
            ret_list_str = ','.join(map(str, ret_list))
            temp = self.quote_assembly_table.get(ret_list_str, QuoteFK=quote_fk)

        return ret_list, temp
    
    def add_operation_to_quote(self, quote_fk):
        ret_list, temp = self.quote_operation_template()
        for data in temp:
            info_dict = dict(zip(ret_list, data))
            info_dict['QuoteFK'] = quote_fk
            self.quote_assembly_table.insert(info_dict)
    
    def create_bom_quote(self, quote_fk, item_fk, quote_assembly_seq_number_fk, sequence_number, order_by,
                            party_fk=None, tool=0, stop_sequence=0, unit_of_measure_set_fk=1, setup_time = 0.00, scrape_rebate = 0.000,
                            part_width = 0.00, part_length = 0.000, parts_required = 1.000, quantity_reqd = 1.000, min_piece_price = 0.00,
                            parts_per_blank_scrap_percentage = 0.000, markup_percentage_1 = 9.999999, piece_weight = 0.000, custom_piece_weight = 0.0000, 
                            piece_cost = 0.0000, piece_price = 0.00000, stock_pieces =0, stock_pieces_scrap_perc = 0.000,
                            calculation_type_fk=17, unattended_operation=0, do_not_use_delivery_schedule=0,
                            vendor_unit=1.00000, grain_direction=0, parts_per_blank=1.000,
                            against_grain=0, double_sided=0, cert_reqd=0, non_amortized_item=0,
                            pull=0, not_include_in_piece_price=0, lock=0, nestable=0,
                            bulk_ship=0, ship_loose=0, customer_supplied_material=0, thickness=0.00):        
        info_dict = {
            "QuoteFK": quote_fk, 
            "ItemFK": item_fk, 
            "PartyFK": party_fk, 
            "UnitOfMeasureSetFK": unit_of_measure_set_fk, 
            "CalculationTypeFK": calculation_type_fk,
            "Tool": tool, 
            "StopSequence": stop_sequence, 
            "SequenceNumber": sequence_number, 
            "QuoteAssemblySeqNumberFK": quote_assembly_seq_number_fk, 
            "UnattendedOperation": unattended_operation,
            "DoNotUseDeliverySchedule": do_not_use_delivery_schedule, 
            "VendorUnit": vendor_unit, 
            "GrainDirection": grain_direction, 
            "PartsPerBlank": parts_per_blank, 
            "AgainstGrain": against_grain,
            "DoubleSided": double_sided, 
            "CertificationsRequired": cert_reqd, 
            "NonAmortizedItem": non_amortized_item, 
            "Pull": pull, 
            "NotIncludeInPiecePrice": not_include_in_piece_price,
            "Lock": lock, 
            "Nestable": nestable, 
            "BulkShip": bulk_ship, 
            "ShipLoose": ship_loose, 
            "CustomerSuppliedMaterial": customer_supplied_material, 
            "OrderBy": order_by, 
            "SetupTime": setup_time, 
            "ScrapRebate": scrape_rebate, 
            "PartWidth": part_width,
            "PartLength": part_length,
            "PartsRequired": parts_required, 
            "QuantityRequired": quantity_reqd, 
            "MinimumPiecePrice": min_piece_price,
            "PartsPerBlankScrapPercentage": parts_per_blank_scrap_percentage, 
            "MarkupPercentage1": markup_percentage_1,
            "PieceWeight": piece_weight, 
            "CustomPieceWeight": custom_piece_weight, 
            "PieceCost": piece_cost, 
            "PiecePrice": piece_price, 
            "StockPieces": stock_pieces,
            "StockPiecesScrapPercentage": stock_pieces_scrap_perc,
            "Thickness": thickness, 
        }
        self.quote_assembly_table.insert(info_dict)
    
    def create_rfq_line_item(self, item_fk: int, request_for_quote_fk: int, line_reference_number: int, quote_fk: int, price_type_fk=3, unit_of_measure_set_fk=1, quantity=None):
        info_dict = {
            "ItemFK" : item_fk, 
            "RequestForQuoteFK": request_for_quote_fk, 
            "LineReferenceNumber": line_reference_number, 
            "QuoteFK": quote_fk, 
            "Quantity": quantity, 
            "PriceTypeFK": price_type_fk, 
            "UnitOfMeasureSetFK": unit_of_measure_set_fk
        }
        pk = self.rfq_line_table.insert(info_dict)
        return pk
    
    def rfq_line_qty(self, rfq_line_fk, quantity, delivery = 1, price_type_fk = 3):
        info_dict = {
            "RequestForQuoteLineFK": rfq_line_fk, 
            "PriceTypeFK": price_type_fk, 
            "Quantity": quantity, 
            "Delivery": delivery,
        }
        self.rfq_line_qty_table.insert(info_dict)
    
    def create_assy_quote(self, quote_to_be_added, quotefk, qty_req = 1):
        info_dict = {
            "QuoteFK": quotefk, 
            "ItemQuoteFK": quote_to_be_added, 
            "SequenceNumber": 1, 
            "Pull": 0, 
            "Lock": 0, 
            "OrderBy": 1, 
            "QuantityRequired": qty_req
        }
        pk = self.quote_assembly_table.insert(info_dict)
        ret_list, temp = self.quote_operation_template(quote_fk=quote_to_be_added)
        for data in temp:
            info_dict1 = dict(zip(ret_list, data))
            info_dict1['QuoteFK'] = quotefk
            info_dict1['ParentQuoteAssemblyFK'] = pk
            info_dict1['ParentQuoteFK'] = quote_to_be_added
            self.quote_assembly_table.insert(info_dict1)
    
    def insert_part_details_in_item(self, item_pk, part_number, values, item_type = None):
        if item_type == 'Material':
            po_comment = f" Dimensions (L x W x T): {values[7]} x {values[8]} x {values[1]}"
            self.item_table.update(item_pk, StockLength=values[7], Thickness=values[1], StockWidth=values[8], Weight=values[3], PartLength=values[0], PartWidth=values[2], PurchaseOrderComment=po_comment, ManufacturedItem=0, Purchase=1, ShipLoose=0, BulkShip=0)
        else:
            self.item_table.update(item_pk, StockLength=values[7], Thickness=values[1], StockWidth=values[8], Weight=values[3], DrawingNumber=values[4], DrawingRevision=values[5], Revision=values[6], PartLength=values[0], PartWidth=values[2], VendorPartNumber=part_number)

    def create_buyer(self, info_dict, customer_pk):
        """ """
        buyer_pk = self.party_table.insert(info_dict)
        my_dict = {
            "PartyFK": customer_pk,
            "BuyerFK": buyer_pk,
        }
        self.party_buyer_table.insert(my_dict)
        return buyer_pk
    
    def create_quote_assembly_formula_variable(self, quote_pk):
        temp = self.quote_assembly_table.get("QuoteAssemblyPK", "SetupFormulaFK", "RunFormulaFK", "OperationFK", QuoteFK=quote_pk)
        for a,b,c,d in temp:
            if d:
                dict_1 = {
                    "QuoteAssemblyFK" : a,
                    "OperationFormulaVariableFK": b,
                    "FormulaType": 0,
                    "VariableValue": 0.000,
                }
                dict_2 = {
                    "QuoteAssemblyFK" : a,
                    "OperationFormulaVariableFK": c,
                    "FormulaType": 1,
                    "VariableValue": 0.000,
                }
                self.quote_assembly_formula_variable_table.insert(dict_1)
                self.quote_assembly_formula_variable_table.insert(dict_2)
    
    def create_item(self, part_number, partyfk, stock_width, stock_length, thickness, weight, item_type_fk=1, mps_item=1, purchase=1, forecast_on_mrp=1, mps_on_mrp=1, service_item=1, unit_of_measure_set_fk=1, vendor_unit=1.0, manufactured_item=0, calculation_type_fk=1, inventoriable=1, purchase_order_comment=None,  description=None, comment=None, bulk_ship=1, ship_loose=1, cert_reqd_by_supplier=0, can_not_create_work_order=0, can_not_invoice=0, general_ledger_account_fk=100, purchase_account_fk=116, cogs_acc_fk=116):
        inventory_info_dict = {
                "QuantityOnHand": 0.000,
            }
        item_inventory_pk = self.item_inventory_table.insert(inventory_info_dict)
        item_info_dict = {
            "ItemInventoryFK" : item_inventory_pk, 
            "PartyFK" : partyfk,
            "PartNumber": part_number, 
            "ItemTypeFK": item_type_fk, 
            "Description" : description, 
            "Comment": comment, 
            "MPSItem": mps_item,
            "Purchase": purchase, 
            "ForecastOnMRP": forecast_on_mrp, 
            "MPSOnMRP": mps_on_mrp, 
            "ServiceItem": service_item, 
            "PurchaseOrderComment": purchase_order_comment, 
            "UnitOfMeasureSetFK": unit_of_measure_set_fk,
            "VendorUnit": vendor_unit, 
            "ManufacturedItem": manufactured_item, 
            "CalculationTypeFK": calculation_type_fk, 
            "Inventoriable": inventoriable, 
            "BulkShip": bulk_ship, 
            "ShipLoose": ship_loose,
            "CertificationsRequiredBySupplier": cert_reqd_by_supplier,
            "CanNotCreateWorkOrder": can_not_create_work_order,
            "CanNotInvoice": can_not_invoice,
            "GeneralLedgerAccountFK": general_ledger_account_fk,
            "PurchaseGeneralLedgerAccountFK" : purchase_account_fk,
            "SalesCogsAccountFK": cogs_acc_fk,
            "StockWidth": stock_width,
            "StockLength": stock_length,
            "Weight": weight,
            "Thickness": thickness,
        }
        item_pk = self.item_table.insert(item_info_dict)
        return item_pk
    
    # Update: May10
    def create_router(self, item_fk, part_number, division_fk=1, router_status_fk=2, router_type=0, default_router=1):
        router_dict = {
            "ItemFK": item_fk,
            "RouterStatusFK": router_status_fk,
            "RouterType": router_type,
            "DefaultRouter": default_router,
            "PartNumber": part_number,
            "DivisionFK": division_fk,
        }
        router_pk = self.router_table.insert(router_dict)
        return router_pk
    
    def create_router_work_center(self, item_fk, router_fk, order_by, unit_of_measure_set_fk=1, sequence_number=1, parts_per_blank=1.00, parts_reqd = 1.000, qty_reqd= 1.000, qty_per_inv = 1, min_per_part=0, vend_unit= 1.00, setup_time=0.00):
        router_work_center_dict = {
            "ItemFK": item_fk,
            "RouterFK": router_fk,
            "OrderBy": order_by,
            "UnitOfMeasureSetFK": unit_of_measure_set_fk,
            "SequenceNumber": sequence_number,
            "PartsPerBlank": parts_per_blank,
            "PartsRequired": parts_reqd,
            "QuantityRequired": qty_reqd,
            "QuantityPerInverse": qty_per_inv,
            "MinutesPerPart": min_per_part,
            "VendorUnit": vend_unit,
            "SetupTime": setup_time,
        }
        self.router_work_center_table.insert(router_work_center_dict)
