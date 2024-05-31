import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import Calendar
from threading import Thread
from src.general_class import TableManger
from src.mie_trak import MieTrak
from src.helper import create_dict_from_excel, transfer_file_to_folder, pk_info_dict, check_and_create_tooling
import os
import datetime

class LoadingScreen(tk.Toplevel):
    def __init__(self, master, max_progress):
        super().__init__(master)
        self.title("Generating RFQ")
        self.geometry("300x100")
        self.protocol("WM_DELETE_WINDOW", self.disable_close_button)
        self.attributes("-topmost", True)  # Ensure loading screen stays on top
        self.grab_set()
        self.progressbar = ttk.Progressbar(self, orient="horizontal", length=200, mode="determinate", maximum=max_progress)
        self.progressbar.pack(pady=10)
    
    def set_progress(self, value):
        self.progressbar["value"] = value
        if value >= self.progressbar["maximum"]:
            self.destroy()
    
    def disable_close_button(self):
        pass

class AddBuyerScreen(tk.Toplevel):
    def __init__(self, master, party_pk, party_name):
        super().__init__(master)
        self.title(f"Add Buyer for Customer: {party_name}, PartyPK: {party_pk}")
        self.geometry("500x300")
        self.attributes("-topmost", True)
        self.grab_set()
        tk.Label(self, text="Name: ").grid(row=0, column=0)
        self.buyer_name_box = tk.Entry(self, width=20)
        self.buyer_name_box.grid(row=0, column=1)
        tk.Label(self, text="Short Name: ").grid(row=1, column=0)
        self.short_name_box = tk.Entry(self, width=20)
        self.short_name_box.grid(row=1, column=1)
        tk.Label(self, text="Email: ").grid(row=2, column=0)
        self.email_id_box = tk.Entry(self, width=20)
        self.email_id_box.grid(row=2, column=1)
        tk.Label(self, text="Phone Number: ").grid(row=3, column=0)
        self.phone_number_box = tk.Entry(self, width=20)
        self.phone_number_box.grid(row=3,column=1)
        tk.Label(self, text="Title: ").grid(row=4, column=0)
        self.title_box = tk.Entry(self, width=20)
        self.title_box.grid(row=4, column=1)
        save_button = tk.Button(self, text="Save", command=lambda: self.save_buyer_info(party_pk))
        save_button.grid(row=5, column=1)
    
    def save_buyer_info(self, party_pk):
        if self.buyer_name_box.get():
            buyer_info_dict = {
                "Name" : self.buyer_name_box.get(),
                "Email": self.email_id_box.get(),
                "Phone": self.phone_number_box.get(),
                "ShortName": self.short_name_box.get(),
                "Title": self.title_box.get(),
                "Buyer": 1,
                "HardwareCertificationFK": 1,
                "MaterialCertificationFK": 1,
                "OutsideProcessingCertificationFK": 1,
                "QualityLevelFK": 2,
                "KeepDocumentOnFile": 1,
                "FirstArticleFK": 1
            }
            buyer_pk = MieTrak().create_buyer(buyer_info_dict, party_pk)
            messagebox.showinfo("Success", f"Buyer created successfully! BuyerPK: {buyer_pk}")
            self.destroy()
        else:
            messagebox.showerror("ERROR", "Please Enter Name")

class RfqGen(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RFQGen")
        self.geometry("500x425")
        self.data_base_conn = MieTrak()
        self.customer_names = self.data_base_conn.get_customer_data(names=True)
        self.quote_assembly_table = TableManger("QuoteAssembly")
        self.make_combobox()
    
    def filter_combobox(self, event):
        current_text = self.customer_select_box.get().lower()
        self.customer_select_box['values'] = ()
        filtered_values = [name for name in self.customer_names if name.lower().startswith(current_text)]
        self.filtered_indices = [idx for idx, name in enumerate(self.customer_names) if name in filtered_values]
        self.customer_select_box['values'] = filtered_values
    
    def filter_buyer_box(self, event):
        current_text = self.buyer_select_box.get().lower()
        self.buyer_select_box['values'] = ()
        filtered_values = [name for name in list(self.buyer_dict.keys()) if name.lower().startswith(current_text)]
        self.buyer_select_box['values'] = filtered_values
    
    def make_combobox(self):
        tk.Label(self, text="Select Customer: ").grid(row=0, column=0)
        self.customer_select_box = ttk.Combobox(self, values=self.customer_names, state="normal")
        self.customer_select_box.grid(row=1, column=0)
        
        tk.Label(self, text="Selected Customer/ Buyer Info: ").grid(row=2, column=0)
        self.customer_info_text = tk.Text(self, height=4, width=30)
        self.customer_info_text.grid(row=3, column=0)

        self.filtered_indices = []
        self.customer_select_box.bind('<KeyRelease>', self.filter_combobox)
        self.customer_select_box.bind("<<ComboboxSelected>>", self.update_customer_info)

        tk.Label(self, text="Select Buyer: ").grid(row=0, column=1)
        self.buyer_select_box = ttk.Combobox(self, state="normal")
        self.buyer_select_box.grid(row=1, column=1)

        self.buyer_select_box.bind('<KeyRelease>', self.filter_buyer_box)

        add_buyer_button = tk.Button(self, text="ADD Buyer", command=self.open_add_buyer_screen)
        add_buyer_button.grid(row=2, column=1)

        self.buyer_select_box.bind("<<ComboboxSelected>>", self.update_buyer_info)

        tk.Label(self, text="Enter Customer RFQ Number: ").grid(row=4, column=0)
        self.rfq_number_text = tk.Entry(self, width=20)
        self.rfq_number_text.grid(row=5, column=0)

        tk.Label(self, text="Parts Requested File:").grid(row=6, column=0)
        self.file_path_PR_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PR_entry.grid(row=7, column=0)

        browse_button_1 = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("Excel files", self.file_path_PR_entry))
        browse_button_1.grid(row=8, column=0)

        tk.Label(self, text="Part Lists File (PL):").grid(row=9, column=0)
        self.file_path_PL_entry = tk.Listbox(self, height=2, width=50)
        self.file_path_PL_entry.grid(row=10, column=0)

        browse_button_part_list = tk.Button(self, text="Browse Files", command=lambda: self.browse_files_parts_requested("All files", self.file_path_PL_entry))
        browse_button_part_list.grid(row=11, column=0)

        self.itar_restricted_var = tk.BooleanVar()
        self.itar_restricted_checkbox = tk.Checkbutton(self, text="ITAR RESTRICTED", variable=self.itar_restricted_var)
        self.itar_restricted_checkbox.grid(row=12, column=0)

        generate_button = tk.Button(self, text="Generate RFQ", command=self.generate_rfq_with_loading_screen)
        generate_button.grid(row=13, column=0)

        add_item_button = tk.Button(self, text="ADD Item", command=self.add_item)
        add_item_button.grid(row=14, column=0)

        tk.Label(self, text="Enter Inquiry Date (MM/DD/YYYY): ").grid(row=3, column=1)
        self.inquiry_date_box = tk.Entry(self, width=20)
        self.inquiry_date_box.grid(row=4, column=1)
        cal_button = tk.Button(self, text='Cal', command=self.open_calendar)
        cal_button.grid(row=5, column=1)

        tk.Label(self, text="Enter Due Date (MM/DD/YYYY): ").grid(row=6, column=1)
        self.due_date_box = tk.Entry(self, width=20)
        self.due_date_box.grid(row=7, column=1)
        cal_due_button = tk.Button(self, text='Cal', command=self.open_due_calendar)
        cal_due_button.grid(row=8, column=1)
    
    def open_calendar(self):
        top = tk.Toplevel(self)
        top.grab_set()
        self.inq_cal = Calendar(top, selectmode="day", date_pattern="mm/dd/y")
        self.inq_cal.pack(padx=20, pady=20)

        top.bind('<Double-1>', self.get_selected_inquiry_date)

        btn = tk.Button(top, text="Get Selected Date", command=self.get_selected_inquiry_date)
        btn.pack(pady=10)
    
    def get_selected_inquiry_date(self, event=None):
        selected_date = self.inq_cal.get_date()
        self.inquiry_date_box.delete(0, tk.END)
        self.inquiry_date_box.insert(tk.END, selected_date)
        self.inq_cal.master.destroy()
    
    def open_due_calendar(self):
        top = tk.Toplevel(self)
        top.grab_set()
        self.due_cal = Calendar(top, selectmode="day", date_pattern="mm/dd/y")
        self.due_cal.pack(padx=20, pady=20)

        top.bind('<Double-1>', self.get_selected_due_date)

        btn = tk.Button(top, text="Get Selected Date", command=self.get_selected_due_date)
        btn.pack(pady=10)
    
    def get_selected_due_date(self, event=None):
        selected_date = self.due_cal.get_date()
        self.due_date_box.delete(0, tk.END)
        self.due_date_box.insert(tk.END, selected_date)
        self.due_cal.master.destroy()
    
    def generate_rfq_with_loading_screen(self):
        self.loading_screen = LoadingScreen(self, max_progress=100)
        Thread(target=self.generate_rfq, args=(self.loading_screen,)).start()  # Start RFQ generation in a separate thread
    
    def open_add_buyer_screen(self):
        if self.customer_select_box.get():
            party_pk = self.party_pk
            name = self.customer_select_box.get()
            AddBuyerScreen(self, party_pk, name)
        else:
            messagebox.showerror("ERROR", "First Select Customer")

    def update_customer_info(self, event=None):
        self.customer_info_text.delete(1.0, tk.END)
        self.buyer_select_box.set("")
        if self.filtered_indices:
            current_index = self.customer_select_box.current()
            selected_customer_index = self.filtered_indices[current_index]
            short_name, email, party_pk = self.data_base_conn.get_customer_data(selected_customer_index=selected_customer_index)
            self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")
            self.party_pk = party_pk
        else:
            short_name, email, party_pk = self.data_base_conn.get_customer_data(selected_customer_index=self.customer_select_box.current())
            self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")
            self.party_pk = party_pk
        
        self.update_buyer_combobox()
    
    def update_buyer_info(self, event=None):
        self.customer_info_text.delete(1.0, tk.END)
        buyer = self.buyer_select_box.get()
        buyer_fk = self.buyer_dict[buyer]
        short_name, email = self.data_base_conn.get_buyer_info(buyer_fk)
        self.customer_info_text.insert(tk.END, f"Name: {short_name}\nEmail: {email}")

    def browse_files_parts_requested(self, filetype: str, list_box):
        """ Browse button for Part requested section, filetype only accepts -> "All files", "Excel files" """
        if filetype == "Excel files":
            param = (filetype, "*.xlsx;*.xls")
        else:
            param = (filetype, "*.*")

        try:
            self.filepaths = [filepath for filepath in filedialog.askopenfilenames(title="Select Files", filetypes=(param,))]

            # entering all file paths in the listbox
            list_box.delete(0, tk.END)
            for path in self.filepaths:
                list_box.insert(0, path)

        except FileNotFoundError as e:
            print(f"Error during file browse: {e}")
            messagebox.showerror("File Browse Error", "An error occurred during file selection. Please try again.")
    
    def update_buyer_combobox(self, event=None):
        self.buyer_dict = self.data_base_conn.get_buyer_data(self.party_pk)
        self.buyer_select_box['values'] = list(self.buyer_dict.keys())
    
    def generate_rfq(self, loading_screen):
        """ Main function for generating RFQ """
        if self.customer_select_box.get() and self.file_path_PR_entry.get(0):
            party_pk = self.party_pk
            billing_details, state, country = self.data_base_conn.get_address(party_pk)
            customer_rfq_number = self.rfq_number_text.get()
            inquiry_date = self.inquiry_date_box.get()
            due_date = self.due_date_box.get()
            current_date = datetime.date.today()
            if current_date:
                formatted_date = current_date.strftime("%m-%d-%Y")
                current_date_formatted = f"{formatted_date} 12:00:00 AM"
            else:
                current_date_formatted = None
            if inquiry_date:
                inq_date = f"{inquiry_date} 12:00:00 AM"
            else:
                inq_date = None
            
            if due_date:
                due_date_formated = f"{due_date} 12:00:00 AM"
            else:
                due_date_formated = None
            if self.buyer_select_box.get(): 
                buyer_fk = self.buyer_dict[self.buyer_select_box.get()]
            else: 
                buyer_fk = None                
            rfq_pk = self.data_base_conn.insert_into_rfq(party_pk, billing_details, state, country, customer_rfq_number=customer_rfq_number, buyer_fk= buyer_fk, inquiry_date=inq_date, due_date=due_date_formated, create_date=current_date_formatted)
            path_dict = {} #dictionary with file path as key and the pk of the document group
            user_selected_file_paths = list(self.file_path_PR_entry.get(0, tk.END) + self.file_path_PL_entry.get(0, tk.END)) #making a list of file paths that user uploaded
            i = 1
            y = 1
            count = 1
            info_dict = create_dict_from_excel(self.file_path_PR_entry.get(0, tk.END)[0])
            my_dict = pk_info_dict(info_dict)
            restricted = False
            quote_pk_dict = {}
            loading_screen.set_progress(10)
            ct=20
            for key, value in info_dict.items():
                if value[13] is None:
                    if self.itar_restricted_var.get(): # checking if the user clicked on Restricted box or not and based on that destination path is decided
                        destination_path = rf'y:\PDM\Restricted\{self.customer_select_box.get()}\{key}'
                        restricted = True
                    else:
                        destination_path = rf'y:\PDM\Non-restricted\{self.customer_select_box.get()}\{key}'
                    
                    for file in user_selected_file_paths:
                        # folder is get or created and file is copied to this folder
                        file_path_to_add_to_rfq = transfer_file_to_folder(destination_path, file)
                        path = file_path_to_add_to_rfq.lower()
                        if "_pl_" in path or "spdl" in path or "psdl" in path or "pl" in os.path.basename(path):
                            path_dict[file_path_to_add_to_rfq] = 26
                        elif "dwg" in path or "drw" in path:
                            path_dict[file_path_to_add_to_rfq] = 27
                        elif "step" in path or "stp" in path:
                            path_dict[file_path_to_add_to_rfq] = 30
                        elif "zsp" in path or "speco" in path:
                            path_dict[file_path_to_add_to_rfq] = 33
                        elif ".cat" in path:
                            path_dict[file_path_to_add_to_rfq] = 16
                        else:
                            path_dict[file_path_to_add_to_rfq] = None
                    
                    for file, pk in path_dict.items():
                        if count==1:
                            if restricted:
                                self.data_base_conn.upload_documents(file, rfq_fk=rfq_pk, document_type_fk=6, secure_document=1, document_group_pk=pk)
                            else:
                                self.data_base_conn.upload_documents(file, rfq_fk=rfq_pk, document_type_fk=6, document_group_pk=pk)
                    count+=1

                    item_pk = self.data_base_conn.get_or_create_item(key, description=value[0], purchase=0, service_item=0, manufactured_item=1)
                    matching_paths = {path:pk for path,pk in path_dict.items() if key in path}
                    for url, pk in matching_paths.items():
                            if restricted:
                                self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, secure_document=1, document_group_pk=pk)
                            else:
                                self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, document_group_pk=pk)
                    
                    quote_pk = self.data_base_conn.create_quote(party_pk, item_pk, 0, key)
                    quote_pk_dict[key] = quote_pk
                    self.data_base_conn.add_operation_to_quote(quote_pk)

                    a = [6,21,22]
                    quote_assembly_fk = []

                    for x in a:
                        quote_assembly_pk = self.quote_assembly_table.get("QuoteAssemblyPK", QuoteFK=quote_pk, SequenceNumber=x)
                        quote_assembly_fk.append(quote_assembly_pk[0][0])
                    
                    if key in my_dict:
                        dict_values = my_dict[key]
                        for j, k, l in zip(dict_values, quote_assembly_fk, a):   # noqa: E741
                            if j is not None and l == 6:
                                self.data_base_conn.create_bom_quote(quote_pk, j, k, l, y, part_length=value[14], part_width=value[15], thickness=value[16])
                                y+=1
                            elif j is not None:
                                self.data_base_conn.create_bom_quote(quote_pk, j, k, l, y, part_length=value[1], part_width=value[3], thickness=value[2])
                                y+=1
                        # Update: May10
                        if dict_values[2]:
                            op_finish_pk = dict_values[2]
                            op_part_number = f"{key} - OP Finish"
                            finish_description = value[6]
                            self.create_finish_router(finish_description, op_finish_pk, op_part_number)
                    
                    if value[12] is None:
                        rfq_line_pk = self.data_base_conn.create_rfq_line_item(item_pk, rfq_pk, i, quote_pk, quantity=value[10])
                        i+=1
                        self.data_base_conn.rfq_line_qty(rfq_line_pk, value[10])
                    else:
                        part_num = value[12]
                        fk = quote_pk_dict.get(part_num)
                        self.data_base_conn.create_assy_quote(quote_pk, fk, qty_req=value[10])
                    
                    if key in info_dict:
                        dict_values = [value[1], value[2], value[3], value[4], value[8], value[9], value[11], value[14], value[15], value[16]]
                        self.data_base_conn.insert_part_details_in_item(item_pk, key, dict_values)
                        pk_value = my_dict[key]
                        for j in pk_value[1:]:
                            if j:
                                self.data_base_conn.insert_part_details_in_item(j, key, dict_values)
                                for url, pk in matching_paths.items():
                                    if restricted:
                                        self.data_base_conn.upload_documents(url, item_fk=j, document_type_fk=2, secure_document=1, document_group_pk=pk, print_with_purchase_order=1)
                                    else:
                                        self.data_base_conn.upload_documents(url, item_fk=j, document_type_fk=2, document_group_pk=pk, print_with_purchase_order=1)
                        if pk_value[0]:
                            self.data_base_conn.insert_part_details_in_item(pk_value[0], key, dict_values, item_type='Material')

                else:
                    part_num = value[12]
                    fk = quote_pk_dict.get(part_num)
                    if value[13] == "Hardware":
                        quote_assembly_pk = self.quote_assembly_table.get("QuoteAssemblyPK", QuoteFK=fk, SequenceNumber=24)
                        # item_fk = self.data_base_conn.get_or_create_item(key, item_type_fk=3, description=value[0], calculation_type_fk=12, purchase_account_fk=130, cogs_acc_fk=130, mps_item=0, forecast_on_mrp=0,mps_on_mrp=0,service_item=0,ship_loose=0,bulk_ship=0)
                        item_fk = check_and_create_tooling(value[0])
                        self.data_base_conn.create_bom_quote(fk, item_fk, quote_assembly_pk[0][0], 24, y)
                        y+=1
                    elif value[13] == "Tooling":
                        quote_assembly_pk = self.quote_assembly_table.get("QuoteAssemblyPK", QuoteFK=fk, SequenceNumber=8)
                        item_fk = self.data_base_conn.get_or_create_item(key, description=value[0], item_type_fk=7, mps_item=0, purchase=0, forecast_on_mrp=0, mps_on_mrp=0, service_item=0, ship_loose=0, bulk_ship=0, can_not_create_work_order=1, can_not_invoice=1, manufactured_item=1)
                        self.data_base_conn.create_bom_quote(fk, item_fk, quote_assembly_pk[0][0], 8, y)
                        y+=1
                loading_screen.set_progress(ct)
                if ct<90:
                    ct+=10
            for value in quote_pk_dict.values():
                self.data_base_conn.create_quote_assembly_formula_variable(value)
            loading_screen.set_progress(100)
            messagebox.showinfo("Success", f"RFQ generated successfully! RFQ Number: {rfq_pk}")
            self.customer_select_box.set("")
            self.buyer_select_box.set("")
            self.customer_info_text.delete(1.0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.file_path_PR_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)
            self.inquiry_date_box.delete(0, tk.END)
            self.due_date_box.delete(0, tk.END)
        else:
            self.loading_screen.destroy()
            messagebox.showerror("ERROR", "Select Customer/ Upload Parts Requested File")
            self.customer_select_box.set("")
            self.buyer_select_box.set("")
            self.customer_info_text.delete(1.0, tk.END)
            self.file_path_PR_entry.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)
            self.inquiry_date_box.delete(0, tk.END)
            self.due_date_box.delete(0, tk.END)
    
    # Update: May10
    def create_finish_router(self, finish_description, item_fin_pk, part_num):
        finish_code = finish_description.split('\n')
        finish_pks = []
        i=1
        if finish_code:
            for code in finish_code:
                finish_codes_pk = self.data_base_conn.get_or_create_item(part_number=code, description=code, inventoriable=0, item_type_fk=5, cert_reqd_by_supplier=1, can_not_create_work_order=1, can_not_invoice=1, purchase_account_fk=125, cogs_acc_fk=125, calculation_type_fk=17)
                finish_pks.append(finish_codes_pk)
        router_pk = self.data_base_conn.create_router(item_fin_pk, part_num)
        for pk in finish_pks:
            self.data_base_conn.create_router_work_center(pk, router_pk, i)
            i+=1

    def add_item(self):
        if self.file_path_PR_entry.get(0):
            if self.customer_select_box.get():
                party_pk = self.party_pk
            else:
                party_pk = None
            user_selected_file_paths = list(self.file_path_PR_entry.get(0, tk.END) + self.file_path_PL_entry.get(0, tk.END))
            info_dict = create_dict_from_excel(self.file_path_PR_entry.get(0, tk.END)[0])
            path_dict = {}
            restricted = False
            for key, value in info_dict.items():
                if self.customer_select_box.get():
                    if self.itar_restricted_var.get(): # checking if the user clicked on Restricted box or not and based on that destination path is decided
                        destination_path = rf'y:\PDM\Restricted\{self.customer_select_box.get()}\{key}'
                        restricted = True
                    else:
                        destination_path = rf'y:\PDM\Non-restricted\{self.customer_select_box.get()}\{key}'
                    
                    for file in user_selected_file_paths:
                        # folder is get or created and file is copied to this folder
                        file_path_to_add_to_rfq = transfer_file_to_folder(destination_path, file)
                        path = file_path_to_add_to_rfq.lower()
                        if "_pl_" in path or "spdl" in path or "psdl" in path or "pl" in os.path.basename(path):
                            path_dict[file_path_to_add_to_rfq] = 26
                        elif "dwg" in path or "drw" in path:
                            path_dict[file_path_to_add_to_rfq] = 27
                        elif "step" in path or "stp" in path:
                            path_dict[file_path_to_add_to_rfq] = 30
                        elif "zsp" in path or "speco" in path:
                            path_dict[file_path_to_add_to_rfq] = 33
                        elif ".cat" in path:
                            path_dict[file_path_to_add_to_rfq] = 16
                        elif "prt" in path:
                            path_dict[file_path_to_add_to_rfq] = 17
                        elif "lwg" in path:
                            path_dict[file_path_to_add_to_rfq] = 29
                        else:
                            path_dict[file_path_to_add_to_rfq] = None
                if value[13] == "Hardware":
                    item_pk = self.data_base_conn.create_item(key, party_pk, value[15] , value[14], value[2], value[4], item_type_fk=3, description=value[0])
                elif value[13] == "Tooling":
                    item_pk = self.data_base_conn.create_item(key, party_pk, value[15] , value[14], value[2], value[4], item_type_fk=7, description=value[0], purchase=0, forecast_on_mrp=0, can_not_create_work_order=1, can_not_invoice=1, manufactured_item=1, mps_item=0, mps_on_mrp=0, service_item=0, bulk_ship=0, ship_loose=0)
                else:
                    item_pk = self.data_base_conn.create_item(key, party_pk, value[15] , value[14], value[2], value[4], description=value[0])

                if self.customer_select_box.get():
                    matching_paths = {path:pk for path,pk in path_dict.items() if key in path}
                    for url, pk in matching_paths.items():
                        if restricted:
                            self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, secure_document=1, document_group_pk=pk)
                        else:
                            self.data_base_conn.upload_documents(url, item_fk=item_pk, document_type_fk=2, document_group_pk=pk)
            messagebox.showinfo("Success", "Item added successfully!")
            self.customer_select_box.set("")
            self.buyer_select_box.set("")
            self.customer_info_text.delete(1.0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.file_path_PR_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)
        else:
            messagebox.showerror("ERROR", "Upload Parts to be added File")
            self.customer_select_box.set("")
            self.buyer_select_box.set("")
            self.customer_info_text.delete(1.0, tk.END)
            self.file_path_PR_entry.delete(0, tk.END)
            self.file_path_PL_entry.delete(0, tk.END)
            self.rfq_number_text.delete(0, tk.END)

if __name__ == "__main__":
    r = RfqGen()
    r.mainloop()